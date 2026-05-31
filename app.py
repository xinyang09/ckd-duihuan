import os
import random
import re
import secrets
from datetime import datetime, timedelta, timezone

import pymysql
import requests
from flask import Flask, jsonify, redirect, render_template, request, session, url_for


DEFAULT_CLIENT_ID = os.getenv("OUTLOOK_CLIENT_ID", "")
TOKEN_URL = os.getenv(
    "OUTLOOK_TOKEN_URL",
    "https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
)
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
TABLE_NAME = os.getenv("CDK_TABLE_NAME", "cdk_accounts")
SMS_PHONE_TABLE = os.getenv("SMS_PHONE_TABLE", "sms_phone_pool")
SMS_CDK_TABLE = os.getenv("SMS_CDK_TABLE", "sms_cdks")
APP_PORT = int(os.getenv("APP_PORT", "5090"))
RECENT_MINUTES = int(os.getenv("RECENT_MINUTES", "30"))

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", secrets.token_hex(32))


def env_flag(name, default=False):
    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


def admin_enabled():
    return env_flag("ADMIN_ENABLED", default=False)


def admin_password():
    return os.getenv("ADMIN_PASSWORD", "").strip()


def admin_authenticated():
    return bool(session.get("admin_authed"))


def require_admin_page():
    if not admin_enabled():
        return jsonify({"error": "后台未开启"}), 403
    if not admin_password():
        return jsonify({"error": "后台密码未配置"}), 403
    if not admin_authenticated():
        return redirect(url_for("admin_login_page"))
    return None


def require_admin_api():
    if not admin_enabled():
        return jsonify({"error": "后台未开启"}), 403
    if not admin_password():
        return jsonify({"error": "后台密码未配置"}), 403
    if not admin_authenticated():
        return jsonify({"error": "后台未登录"}), 401
    return None


def load_env_file():
    env_path = os.path.join(os.getcwd(), ".env")
    if not os.path.exists(env_path):
        return
    with open(env_path, "r", encoding="utf-8") as env_file:
        for line in env_file:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip())


def get_db_connection(database_required=True):
    kwargs = {
        "host": os.getenv("MYSQL_HOST", "127.0.0.1"),
        "port": int(os.getenv("MYSQL_PORT", "3306")),
        "user": os.getenv("MYSQL_USER", "root"),
        "password": os.getenv("MYSQL_PASSWORD", ""),
        "charset": "utf8mb4",
        "cursorclass": pymysql.cursors.DictCursor,
        "autocommit": False,
    }
    if database_required:
        kwargs["database"] = os.getenv("MYSQL_DATABASE", "cdk_mail")
    return pymysql.connect(**kwargs)


def init_db():
    db_name = os.getenv("MYSQL_DATABASE", "cdk_mail")

    root_conn = get_db_connection(database_required=False)
    try:
        with root_conn.cursor() as cursor:
            cursor.execute(
                f"CREATE DATABASE IF NOT EXISTS `{db_name}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
            )
        root_conn.commit()
    finally:
        root_conn.close()

    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                CREATE TABLE IF NOT EXISTS `{TABLE_NAME}` (
                    id BIGINT PRIMARY KEY AUTO_INCREMENT,
                    cdk VARCHAR(64) NOT NULL UNIQUE,
                    email VARCHAR(255) NOT NULL,
                    email_password TEXT,
                    client_id VARCHAR(255),
                    gpt_password VARCHAR(255) NOT NULL,
                    redeemed TINYINT(1) DEFAULT 0,
                    refresh_token TEXT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            cursor.execute(
                f"""
                CREATE TABLE IF NOT EXISTS `{SMS_PHONE_TABLE}` (
                    id BIGINT PRIMARY KEY AUTO_INCREMENT,
                    phone VARCHAR(32) NOT NULL UNIQUE,
                    upstream_url TEXT,
                    remark VARCHAR(255),
                    assigned_tinyint TINYINT(1) DEFAULT 0,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            cursor.execute(
                f"""
                CREATE TABLE IF NOT EXISTS `{SMS_CDK_TABLE}` (
                    id BIGINT PRIMARY KEY AUTO_INCREMENT,
                    code VARCHAR(64) NOT NULL UNIQUE,
                    phone VARCHAR(32),
                    status VARCHAR(32) DEFAULT 'unused',
                    batch_name VARCHAR(128),
                    redeemed_at TIMESTAMP NULL DEFAULT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            cursor.execute(f"SHOW COLUMNS FROM `{SMS_PHONE_TABLE}`")
            sms_phone_columns = {row["Field"] for row in cursor.fetchall()}
            if "upstream_url" not in sms_phone_columns:
                cursor.execute(f"ALTER TABLE `{SMS_PHONE_TABLE}` ADD COLUMN upstream_url TEXT")
            if "remark" not in sms_phone_columns:
                cursor.execute(f"ALTER TABLE `{SMS_PHONE_TABLE}` ADD COLUMN remark VARCHAR(255)")
            if "assigned_tinyint" not in sms_phone_columns:
                cursor.execute(f"ALTER TABLE `{SMS_PHONE_TABLE}` ADD COLUMN assigned_tinyint TINYINT(1) DEFAULT 0")
            if "created_at" not in sms_phone_columns:
                cursor.execute(f"ALTER TABLE `{SMS_PHONE_TABLE}` ADD COLUMN created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP")

            cursor.execute(f"SHOW COLUMNS FROM `{SMS_CDK_TABLE}`")
            sms_cdk_columns = {row["Field"] for row in cursor.fetchall()}
            if "phone" not in sms_cdk_columns:
                cursor.execute(f"ALTER TABLE `{SMS_CDK_TABLE}` ADD COLUMN phone VARCHAR(32)")
            if "status" not in sms_cdk_columns:
                cursor.execute(f"ALTER TABLE `{SMS_CDK_TABLE}` ADD COLUMN status VARCHAR(32) DEFAULT 'unused'")
            if "batch_name" not in sms_cdk_columns:
                cursor.execute(f"ALTER TABLE `{SMS_CDK_TABLE}` ADD COLUMN batch_name VARCHAR(128)")
            if "redeemed_at" not in sms_cdk_columns:
                cursor.execute(f"ALTER TABLE `{SMS_CDK_TABLE}` ADD COLUMN redeemed_at TIMESTAMP NULL DEFAULT NULL")
            if "created_at" not in sms_cdk_columns:
                cursor.execute(f"ALTER TABLE `{SMS_CDK_TABLE}` ADD COLUMN created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP")
        conn.commit()
    finally:
        conn.close()


def generate_random_code(prefix="GPT", length=12):
    alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
    chunks = []
    chunk_sizes = [4, 4, max(4, length - 8)]
    for size in chunk_sizes:
        chunks.append("".join(random.choice(alphabet) for _ in range(size)))
    return f"{prefix}-{chunks[0]}-{chunks[1]}-{chunks[2]}"


def generate_unique_cdk(prefix="GPT", length=12, table=TABLE_NAME, field="cdk"):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            for _ in range(50):
                code = generate_random_code(prefix=prefix, length=length)
                cursor.execute(f"SELECT 1 FROM `{table}` WHERE `{field}`=%s", (code,))
                if not cursor.fetchone():
                    return code
        raise RuntimeError("生成随机 CDK 失败，请重试")
    finally:
        conn.close()


def fetch_access_token(refresh_token_val, client_id_val, scope=None):
    session = requests.Session()
    session.trust_env = False
    payload = {
        "client_id": client_id_val,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token_val,
    }
    if scope:
        payload["scope"] = scope
    response = session.post(TOKEN_URL, data=payload, timeout=20)
    response.raise_for_status()
    token_data = response.json()
    access_token = token_data.get("access_token")
    if not access_token:
        raise RuntimeError(token_data.get("error_description", "响应中未找到 access_token"))
    return token_data


def fetch_graph_access_token(refresh_token_val, client_id_val):
    return fetch_access_token(refresh_token_val, client_id_val, "https://graph.microsoft.com/.default")


def graph_headers(access_token):
    return {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}


def extract_verification_code(text):
    patterns = [
        r"(?:verification code|security code|one-time code|one-time password|otp|验证码|校验码|安全代码)[:：\s-]*([A-Za-z0-9]{4,10})",
        r"([A-Za-z0-9]{4,10})[:：\s-]*(?:is your verification code|is your security code|是你的验证码)",
        r"(?:code|验证码)[:：\s-]*([0-9]{4,10})",
        r"\b([0-9]{6})\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1)
    return ""


def is_verification_email(item):
    if item.get("code"):
        return True
    subject = (item.get("subject") or "").lower()
    sender = (item.get("from") or "").lower()
    keywords = ["code", "verification", "verify", "otp", "验证码", "校验码", "安全代码", "一次性密码"]
    return any(keyword in f"{subject} {sender}" for keyword in keywords)


def fetch_graph_messages(access_token, max_items=100):
    session = requests.Session()
    session.trust_env = False
    params = {
        "$top": str(max_items),
        "$select": "id,subject,receivedDateTime,from,bodyPreview",
        "$orderby": "receivedDateTime DESC",
    }
    response = session.get(
        f"{GRAPH_BASE_URL}/me/mailFolders/inbox/messages",
        headers=graph_headers(access_token),
        params=params,
        timeout=20,
    )
    response.raise_for_status()
    return response.json().get("value", [])


def parse_graph_message(message):
    subject_str = message.get("subject") or "(No Subject)"
    received = message.get("receivedDateTime") or ""
    from_obj = (((message.get("from") or {}).get("emailAddress")) or {})
    from_name = from_obj.get("name") or ""
    from_address = from_obj.get("address") or ""
    if from_name and from_address:
        from_str = f"{from_name} <{from_address}>"
    else:
        from_str = from_address or from_name or "(Unknown Sender)"

    date_obj = datetime.fromtimestamp(0, tz=timezone.utc)
    if received:
        try:
            date_obj = datetime.fromisoformat(received.replace("Z", "+00:00"))
        except Exception:
            pass

    body_preview = message.get("bodyPreview") or ""
    return {
        "uid": message.get("id", ""),
        "mailbox": "INBOX",
        "subject": subject_str,
        "from": from_str,
        "date": date_obj.strftime("%Y-%m-%d %H:%M:%S"),
        "code": extract_verification_code(f"{subject_str}\n{body_preview}"),
        "date_obj": date_obj,
    }


def filter_recent_items(items, minutes=RECENT_MINUTES):
    cutoff = datetime.now(timezone.utc) - timedelta(minutes=minutes)
    return [item for item in items if item.get("date_obj") and item["date_obj"] >= cutoff]


def fetch_all_emails(email_address, refresh_token, client_id, max_per_mailbox=100):
    graph_result = fetch_graph_access_token(refresh_token, client_id or DEFAULT_CLIENT_ID)
    access_token = graph_result.get("access_token")
    scope = graph_result.get("scope", "")
    if "https://graph.microsoft.com/Mail.Read" not in scope:
        raise RuntimeError(f"当前 token 不支持 Graph Mail.Read，返回 scope: {scope or '(empty)'}")

    raw_messages = fetch_graph_messages(access_token, max_items=max_per_mailbox)
    found_items = [parse_graph_message(message) for message in raw_messages]
    found_items.sort(key=lambda x: x["date_obj"], reverse=True)
    found_items = filter_recent_items(found_items)
    for item in found_items:
        item.pop("date_obj", None)
    return found_items


def get_account_by_cdk(cdk):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT cdk, email, email_password, client_id, gpt_password, refresh_token, redeemed
                FROM `{TABLE_NAME}`
                WHERE cdk = %s
                """,
                (cdk,),
            )
            row = cursor.fetchone()
            return row
    finally:
        conn.close()


def mark_account_redeemed(cdk):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(f"UPDATE `{TABLE_NAME}` SET redeemed = 1 WHERE cdk = %s", (cdk,))
        conn.commit()
    finally:
        conn.close()


def list_accounts():
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT cdk, email, gpt_password, redeemed, created_at
                FROM `{TABLE_NAME}`
                ORDER BY id DESC
                """
            )
            return cursor.fetchall()
    finally:
        conn.close()


def delete_account(cdk):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(f"DELETE FROM `{TABLE_NAME}` WHERE cdk = %s", (cdk,))
            deleted = cursor.rowcount
        conn.commit()
        return deleted
    finally:
        conn.close()


def import_account_line(raw_line):
    parts = [part.strip() for part in raw_line.strip().split("----")]
    if len(parts) != 4:
        raise ValueError("格式错误，必须为 邮箱----邮箱密码/授权码----client_id----refresh_token")

    email, email_password, client_id, refresh_token = parts
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(f"SELECT cdk FROM `{TABLE_NAME}` WHERE lower(email)=lower(%s)", (email,))
            existing = cursor.fetchone()
            if existing:
                raise ValueError(f"该邮箱已导入，现有 CDK: {existing['cdk']}")

            cdk = generate_unique_cdk(prefix="GPT", length=12, table=TABLE_NAME, field="cdk")
            cursor.execute(
                f"""
                INSERT INTO `{TABLE_NAME}` (cdk, email, email_password, client_id, gpt_password, refresh_token)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (
                    cdk,
                    email,
                    email_password,
                    client_id,
                    os.getenv("DEFAULT_GPT_PASSWORD", "Jijie@123456"),
                    refresh_token,
                ),
            )
        conn.commit()
        return {"cdk": cdk, "email": email}
    finally:
        conn.close()


def import_sms_lines(raw_lines):
    rows = []
    for line in raw_lines.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = [part.strip() for part in line.split("----")]
        if len(parts) < 2:
            raise ValueError("手机号导入格式错误，必须为 手机号----上游链接----备注（备注可选）")
        phone = parts[0]
        upstream_url = parts[1]
        remark = parts[2] if len(parts) > 2 else ""
        rows.append((phone, upstream_url, remark))

    if not rows:
        raise ValueError("没有可导入的手机号")

    conn = get_db_connection()
    inserted = 0
    updated = 0
    generated_codes = []
    try:
        with conn.cursor() as cursor:
            for phone, upstream_url, remark in rows:
                cursor.execute(f"SELECT id, assigned_tinyint FROM `{SMS_PHONE_TABLE}` WHERE phone=%s", (phone,))
                existing = cursor.fetchone()
                if existing:
                    cursor.execute(
                        f"UPDATE `{SMS_PHONE_TABLE}` SET upstream_url=%s, remark=%s WHERE id=%s",
                        (upstream_url, remark, existing["id"]),
                    )
                    updated += 1
                    if existing["assigned_tinyint"]:
                        continue
                    phone_id = existing["id"]
                else:
                    cursor.execute(
                        f"INSERT INTO `{SMS_PHONE_TABLE}` (phone, upstream_url, remark) VALUES (%s, %s, %s)",
                        (phone, upstream_url, remark),
                    )
                    inserted += 1
                    phone_id = cursor.lastrowid

                code = generate_unique_cdk(prefix="SMS", length=10, table=SMS_CDK_TABLE, field="code")
                cursor.execute(
                    f"""
                    INSERT INTO `{SMS_CDK_TABLE}` (code, phone, status, batch_name)
                    VALUES (%s, %s, 'unused', %s)
                    """,
                    (code, phone, "auto-import"),
                )
                cursor.execute(
                    f"UPDATE `{SMS_PHONE_TABLE}` SET assigned_tinyint = 1 WHERE id=%s",
                    (phone_id,),
                )
                generated_codes.append(code)
        conn.commit()
        return {
            "total": len(rows),
            "inserted": inserted,
            "updated": updated,
            "generated": len(generated_codes),
            "codes": generated_codes,
        }
    finally:
        conn.close()


def generate_sms_cdks(count, batch_name, prefix, length):
    conn = get_db_connection()
    generated = []
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"SELECT phone FROM `{SMS_PHONE_TABLE}` WHERE assigned_tinyint = 0 ORDER BY id ASC LIMIT %s",
                (count,),
            )
            phones = cursor.fetchall()
            if not phones:
                raise ValueError("没有可分配的手机号")

            for row in phones:
                phone = row["phone"]
                code = generate_unique_cdk(prefix=prefix or "SMS", length=max(length, 8), table=SMS_CDK_TABLE, field="code")
                cursor.execute(
                    f"""
                    INSERT INTO `{SMS_CDK_TABLE}` (code, phone, status, batch_name)
                    VALUES (%s, %s, 'unused', %s)
                    """,
                    (code, phone, batch_name),
                )
                cursor.execute(
                    f"UPDATE `{SMS_PHONE_TABLE}` SET assigned_tinyint = 1 WHERE phone=%s",
                    (phone,),
                )
                generated.append(code)
        conn.commit()
        return {"count": len(generated), "codes": generated}
    finally:
        conn.close()


def sms_dashboard():
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(f"SELECT COUNT(*) AS count FROM `{SMS_PHONE_TABLE}` WHERE assigned_tinyint = 0")
            available = cursor.fetchone()["count"]
            cursor.execute(f"SELECT COUNT(*) AS count FROM `{SMS_PHONE_TABLE}` WHERE assigned_tinyint = 1")
            assigned = cursor.fetchone()["count"]
            cursor.execute(f"SELECT COUNT(*) AS count FROM `{SMS_CDK_TABLE}` WHERE status = 'unused'")
            unused = cursor.fetchone()["count"]
            cursor.execute(f"SELECT COUNT(*) AS count FROM `{SMS_CDK_TABLE}` WHERE status = 'redeemed'")
            redeemed = cursor.fetchone()["count"]
            cursor.execute(
                f"""
                SELECT code, status, phone, redeemed_at
                FROM `{SMS_CDK_TABLE}`
                ORDER BY id DESC
                LIMIT 20
                """
            )
            recent = cursor.fetchall()
        return {
            "stats": {
                "availablePhones": available,
                "assignedPhones": assigned,
                "unusedCdks": unused,
                "redeemedCdks": redeemed,
            },
            "recentCdks": recent,
        }
    finally:
        conn.close()


def delete_sms_cdk(code):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(f"SELECT phone FROM `{SMS_CDK_TABLE}` WHERE code = %s", (code,))
            row = cursor.fetchone()
            if not row:
                return 0
            phone = row.get("phone")
            cursor.execute(f"DELETE FROM `{SMS_CDK_TABLE}` WHERE code = %s", (code,))
            if phone:
                cursor.execute(
                    f"SELECT COUNT(*) AS count FROM `{SMS_CDK_TABLE}` WHERE phone = %s",
                    (phone,),
                )
                linked = cursor.fetchone()["count"]
                if linked == 0:
                    cursor.execute(
                        f"UPDATE `{SMS_PHONE_TABLE}` SET assigned_tinyint = 0 WHERE phone = %s",
                        (phone,),
                    )
            deleted = cursor.rowcount
        conn.commit()
        return deleted
    finally:
        conn.close()


def list_sms_phones():
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT phone, upstream_url, remark, assigned_tinyint, created_at
                FROM `{SMS_PHONE_TABLE}`
                ORDER BY id DESC
                LIMIT 50
                """
            )
            return cursor.fetchall()
    finally:
        conn.close()


def delete_sms_phone(phone):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"SELECT COUNT(*) AS count FROM `{SMS_CDK_TABLE}` WHERE phone = %s",
                (phone,),
            )
            if cursor.fetchone()["count"] > 0:
                raise ValueError("该手机号下仍有关联 SMS CDK，请先删除对应 SMS CDK")
            cursor.execute(f"DELETE FROM `{SMS_PHONE_TABLE}` WHERE phone = %s", (phone,))
            deleted = cursor.rowcount
        conn.commit()
        return deleted
    finally:
        conn.close()


def get_sms_cdk(code):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT code, phone, status, redeemed_at
                FROM `{SMS_CDK_TABLE}`
                WHERE code = %s
                """,
                (code,),
            )
            return cursor.fetchone()
    finally:
        conn.close()


def redeem_sms_cdk(code):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT code, phone, status, redeemed_at
                FROM `{SMS_CDK_TABLE}`
                WHERE code = %s
                """,
                (code,),
            )
            row = cursor.fetchone()
            if not row:
                raise ValueError("未找到对应 SMS CDK")
            if not row.get("phone"):
                raise ValueError("该 SMS CDK 尚未绑定手机号")

            mode = "existing" if row["status"] == "redeemed" else "redeemed"
            if row["status"] != "redeemed":
                cursor.execute(
                    f"""
                    UPDATE `{SMS_CDK_TABLE}`
                    SET status = 'redeemed', redeemed_at = NOW()
                    WHERE code = %s
                    """,
                    (code,),
                )
        conn.commit()
        row["status"] = "redeemed"
        return row, mode
    finally:
        conn.close()


def build_sms_links(phone):
    base_url = os.getenv("SMS_PUBLIC_BASE_URL", f"http://127.0.0.1:{APP_PORT}")
    safe_phone = phone or ""
    api_url = f"{base_url}/api/sms/code/{safe_phone}"
    pickup_url = f"{base_url}/sms-pickup?phone={safe_phone}"
    return api_url, pickup_url


def fetch_sms_message(phone):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"SELECT upstream_url FROM `{SMS_PHONE_TABLE}` WHERE phone = %s",
                (phone,),
            )
            row = cursor.fetchone()
    finally:
        conn.close()

    if not row:
        raise ValueError("未找到对应手机号")

    upstream_url = (row.get("upstream_url") or "").strip()
    if not upstream_url:
        return {"code": "", "raw": "", "source": ""}

    session = requests.Session()
    session.trust_env = False
    response = session.get(upstream_url, timeout=20)
    response.raise_for_status()
    raw_text = response.text.strip()
    return {
        "code": extract_verification_code(raw_text),
        "raw": raw_text,
        "source": upstream_url,
    }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/admin")
def admin():
    guard = require_admin_page()
    if guard:
        return guard
    return render_template("admin.html")


@app.route("/admin/login")
def admin_login_page():
    if not admin_enabled():
        return jsonify({"error": "后台未开启"}), 403
    if admin_authenticated():
        return redirect(url_for("admin"))
    return render_template("admin_login.html")


@app.route("/admin/logout", methods=["POST"])
def admin_logout():
    session.pop("admin_authed", None)
    return jsonify({"ok": True})


@app.route("/api/admin/login", methods=["POST"])
def admin_login():
    if not admin_enabled():
        return jsonify({"error": "后台未开启"}), 403
    configured_password = admin_password()
    if not configured_password:
        return jsonify({"error": "后台密码未配置"}), 403

    data = request.get_json(silent=True) or {}
    password = (data.get("password") or "").strip()
    if password != configured_password:
        return jsonify({"error": "后台密码错误"}), 401

    session["admin_authed"] = True
    return jsonify({"ok": True})


@app.route("/sms-pickup")
def sms_pickup():
    return render_template("sms_pickup.html")


@app.route("/api/sms/redeem", methods=["POST"])
def sms_redeem():
    data = request.get_json(silent=True) or {}
    cdk = (data.get("cdk") or "").strip()
    if not cdk:
        return jsonify({"error": "SMS CDK 不能为空"}), 400

    try:
        row, mode = redeem_sms_cdk(cdk)
        api_url, pickup_url = build_sms_links(row["phone"])
        return jsonify(
            {
                "cdk": row["code"],
                "phone": row["phone"],
                "pickupToken": row["phone"],
                "apiUrl": api_url,
                "pickupUrl": pickup_url,
                "mode": mode,
            }
        )
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 404
    except Exception as exc:
        return jsonify({"error": f"SMS 兑换失败: {exc}"}), 500


@app.route("/api/sms/code/<phone>", methods=["GET"])
def sms_code(phone):
    phone = (phone or "").strip()
    if not phone:
        return jsonify({"error": "手机号不能为空"}), 400

    try:
        payload = fetch_sms_message(phone)
        return jsonify(payload)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 404
    except Exception as exc:
        return jsonify({"error": f"获取短信失败: {exc}"}), 500


@app.route("/api/account", methods=["POST"])
def account_lookup():
    data = request.get_json(silent=True) or {}
    cdk = (data.get("cdk") or "").strip()
    if not cdk:
        return jsonify({"error": "CDK 不能为空"}), 400

    account = get_account_by_cdk(cdk)
    if not account:
        return jsonify({"error": "未找到对应 CDK"}), 404

    mark_account_redeemed(cdk)
    return jsonify(
        {
            "cdk": account["cdk"],
            "email": account["email"],
            "gpt_password": account["gpt_password"],
        }
    )


@app.route("/api/all-mails", methods=["POST"])
def all_mails():
    data = request.get_json(silent=True) or {}
    cdk = (data.get("cdk") or "").strip()
    if not cdk:
        return jsonify({"error": "CDK 不能为空"}), 400

    account = get_account_by_cdk(cdk)
    if not account:
        return jsonify({"error": "未找到对应 CDK"}), 404

    try:
        mails = fetch_all_emails(
            account["email"],
            account["refresh_token"],
            account.get("client_id") or DEFAULT_CLIENT_ID,
            max_per_mailbox=100,
        )
    except Exception as exc:
        return jsonify({"error": f"获取邮件失败: {exc}"}), 500

    verification_mails = [item for item in mails if is_verification_email(item)]
    slim_items = [
        {
            "mailbox": item["mailbox"],
            "subject": item["subject"],
            "from": item["from"],
            "date": item["date"],
            "code": item["code"] or "",
        }
        for item in verification_mails
    ]
    return jsonify({"items": slim_items, "count": len(slim_items)})


@app.route("/api/latest-code", methods=["POST"])
def latest_code():
    data = request.get_json(silent=True) or {}
    cdk = (data.get("cdk") or "").strip()
    if not cdk:
        return jsonify({"error": "CDK 不能为空"}), 400

    account = get_account_by_cdk(cdk)
    if not account:
        return jsonify({"error": "未找到对应 CDK"}), 404

    try:
        mails = fetch_all_emails(
            account["email"],
            account["refresh_token"],
            account.get("client_id") or DEFAULT_CLIENT_ID,
            max_per_mailbox=100,
        )
    except Exception as exc:
        return jsonify({"error": f"获取验证码失败: {exc}"}), 500

    verification_mails = [item for item in mails if is_verification_email(item)]
    if not verification_mails:
        return jsonify({"error": "未找到最新邮件"}), 404

    latest = verification_mails[0]
    return jsonify(
        {
            "mailbox": latest["mailbox"],
            "subject": latest["subject"],
            "from": latest["from"],
            "date": latest["date"],
            "code": latest["code"] or "",
        }
    )


@app.route("/api/admin/accounts", methods=["GET"])
def admin_accounts():
    guard = require_admin_api()
    if guard:
        return guard
    return jsonify({"items": list_accounts()})


@app.route("/api/admin/accounts/<cdk>", methods=["DELETE"])
def admin_delete_account(cdk):
    guard = require_admin_api()
    if guard:
        return guard
    try:
        deleted = delete_account(cdk)
    except Exception as exc:
        return jsonify({"error": f"删除失败: {exc}"}), 500

    if not deleted:
        return jsonify({"error": "未找到对应账号"}), 404
    return jsonify({"deleted": True, "cdk": cdk})


@app.route("/api/admin/import", methods=["POST"])
def admin_import():
    guard = require_admin_api()
    if guard:
        return guard
    data = request.get_json(silent=True) or {}
    raw_line = (data.get("line") or "").strip()
    if not raw_line:
        return jsonify({"error": "导入内容不能为空"}), 400
    try:
        result = import_account_line(raw_line)
        return jsonify(result)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"导入失败: {exc}"}), 500


@app.route("/api/admin/sms/import", methods=["POST"])
def admin_sms_import():
    guard = require_admin_api()
    if guard:
        return guard
    data = request.get_json(silent=True) or {}
    lines = (data.get("lines") or "").strip()
    if not lines:
        return jsonify({"error": "导入内容不能为空"}), 400
    try:
        result = import_sms_lines(lines)
        return jsonify(result)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"导入失败: {exc}"}), 500


@app.route("/api/admin/sms/cdks/generate", methods=["POST"])
def admin_sms_generate():
    guard = require_admin_api()
    if guard:
        return guard
    data = request.get_json(silent=True) or {}
    count = int(data.get("count") or 0)
    batch_name = (data.get("batchName") or "").strip()
    prefix = (data.get("prefix") or "SMS").strip()
    length = int(data.get("length") or 10)
    if count <= 0:
        return jsonify({"error": "count 必须大于 0"}), 400
    try:
        result = generate_sms_cdks(count, batch_name, prefix, length)
        return jsonify(result)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"生成失败: {exc}"}), 500


@app.route("/api/admin/sms/dashboard", methods=["GET"])
def admin_sms_dashboard():
    guard = require_admin_api()
    if guard:
        return guard
    try:
        return jsonify(sms_dashboard())
    except Exception as exc:
        return jsonify({"error": f"读取 SMS 面板失败: {exc}"}), 500


@app.route("/api/admin/sms/cdks/<code>", methods=["DELETE"])
def admin_delete_sms_cdk(code):
    guard = require_admin_api()
    if guard:
        return guard
    try:
        deleted = delete_sms_cdk(code)
    except Exception as exc:
        return jsonify({"error": f"删除失败: {exc}"}), 500
    if not deleted:
        return jsonify({"error": "未找到对应 SMS CDK"}), 404
    return jsonify({"deleted": True, "code": code})


@app.route("/api/admin/sms/phones", methods=["GET"])
def admin_sms_phones():
    guard = require_admin_api()
    if guard:
        return guard
    try:
        return jsonify({"items": list_sms_phones()})
    except Exception as exc:
        return jsonify({"error": f"读取手机号池失败: {exc}"}), 500


@app.route("/api/admin/sms/phones/<phone>", methods=["DELETE"])
def admin_delete_sms_phone(phone):
    guard = require_admin_api()
    if guard:
        return guard
    try:
        deleted = delete_sms_phone(phone)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"删除失败: {exc}"}), 500
    if not deleted:
        return jsonify({"error": "未找到对应手机号"}), 404
    return jsonify({"deleted": True, "phone": phone})


if __name__ == "__main__":
    load_env_file()
    init_db()
    app.run(host="0.0.0.0", port=APP_PORT, debug=True)
