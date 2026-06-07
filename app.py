import hashlib
import os
import random
import re
import secrets
import time
from datetime import datetime, timedelta, timezone

import pymysql
import pymysql.cursors
import requests
from flask import Flask, jsonify, redirect, render_template, request, session, url_for

try:
    from dbutils.pooled_db import PooledDB as _PooledDB
    _HAS_DBUTILS = True
except ImportError:
    _HAS_DBUTILS = False


DEFAULT_CLIENT_ID = os.getenv("OUTLOOK_CLIENT_ID", "")
TOKEN_URL = os.getenv(
    "OUTLOOK_TOKEN_URL",
    "https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
)
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
TABLE_NAME = os.getenv("CDK_TABLE_NAME", "cdk_accounts")
SMS_PHONE_TABLE = os.getenv("SMS_PHONE_TABLE", "sms_phone_pool")
SMS_CDK_TABLE = os.getenv("SMS_CDK_TABLE", "sms_cdks")
MAILBOX_LIBRARY_TABLE = os.getenv("MAILBOX_LIBRARY_TABLE", "mailbox_library")
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



# ── MySQL 连接池（进程级单例，避免每次请求重新握手）──
def _build_pool():
    return _PooledDB(
        creator=pymysql,
        mincached=3,
        maxcached=10,
        maxconnections=20,
        blocking=True,
        ping=1,
        host=os.getenv("MYSQL_HOST", "127.0.0.1"),
        port=int(os.getenv("MYSQL_PORT", "3306")),
        user=os.getenv("MYSQL_USER", "root"),
        password=os.getenv("MYSQL_PASSWORD", ""),
        database=os.getenv("MYSQL_DATABASE", "cdk_mail"),
        charset="utf8mb4",
        cursorclass=pymysql.cursors.DictCursor,
        autocommit=False,
        connect_timeout=10,
        read_timeout=30,
        write_timeout=30,
    )

_db_pool = None


def get_db_pool():
    global _db_pool
    if _db_pool is None:
        _db_pool = _build_pool()
    return _db_pool


def get_db_connection(database_required=True):
    """从连接池取一条连接；database_required=False 时回退到直连（仅用于 init_db）。"""
    if not database_required:
        conn = pymysql.connect(
            host=os.getenv("MYSQL_HOST", "127.0.0.1"),
            port=int(os.getenv("MYSQL_PORT", "3306")),
            user=os.getenv("MYSQL_USER", "root"),
            password=os.getenv("MYSQL_PASSWORD", ""),
            charset="utf8mb4",
            cursorclass=pymysql.cursors.DictCursor,
            autocommit=False,
            connect_timeout=10,
        )
        return conn
    if _HAS_DBUTILS:
        return get_db_pool().connection()
    # 备用：未安装 dbutils 时直连
    return pymysql.connect(
        host=os.getenv("MYSQL_HOST", "127.0.0.1"),
        port=int(os.getenv("MYSQL_PORT", "3306")),
        user=os.getenv("MYSQL_USER", "root"),
        password=os.getenv("MYSQL_PASSWORD", ""),
        database=os.getenv("MYSQL_DATABASE", "cdk_mail"),
        charset="utf8mb4",
        cursorclass=pymysql.cursors.DictCursor,
        autocommit=False,
        connect_timeout=10,
        read_timeout=30,
        write_timeout=30,
    )


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
                    phone VARCHAR(32),
                    email_password TEXT,
                    client_id VARCHAR(255),
                    gpt_password VARCHAR(255) NOT NULL,
                    redeemed TINYINT(1) DEFAULT 0,
                    sold TINYINT(1) DEFAULT 0,
                    redeemed_at TIMESTAMP NULL DEFAULT NULL,
                    refresh_token TEXT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            cursor.execute(f"SHOW COLUMNS FROM `{TABLE_NAME}`")
            account_columns = {row["Field"] for row in cursor.fetchall()}
            if "phone" not in account_columns:
                cursor.execute(f"ALTER TABLE `{TABLE_NAME}` ADD COLUMN phone VARCHAR(32)")
            if "email_password" not in account_columns:
                cursor.execute(f"ALTER TABLE `{TABLE_NAME}` ADD COLUMN email_password TEXT")
            if "client_id" not in account_columns:
                cursor.execute(f"ALTER TABLE `{TABLE_NAME}` ADD COLUMN client_id VARCHAR(255)")
            if "gpt_password" not in account_columns:
                cursor.execute(f"ALTER TABLE `{TABLE_NAME}` ADD COLUMN gpt_password VARCHAR(255) NOT NULL DEFAULT 'Jijie@123456'")
            if "redeemed" not in account_columns:
                cursor.execute(f"ALTER TABLE `{TABLE_NAME}` ADD COLUMN redeemed TINYINT(1) DEFAULT 0")
            if "sold" not in account_columns:
                cursor.execute(f"ALTER TABLE `{TABLE_NAME}` ADD COLUMN sold TINYINT(1) DEFAULT 0")
            if "redeemed_at" not in account_columns:
                cursor.execute(f"ALTER TABLE `{TABLE_NAME}` ADD COLUMN redeemed_at TIMESTAMP NULL DEFAULT NULL")
            if "refresh_token" not in account_columns:
                cursor.execute(f"ALTER TABLE `{TABLE_NAME}` ADD COLUMN refresh_token TEXT NOT NULL")
            if "created_at" not in account_columns:
                cursor.execute(f"ALTER TABLE `{TABLE_NAME}` ADD COLUMN created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP")
            cursor.execute(
                f"""
                CREATE TABLE IF NOT EXISTS `{SMS_PHONE_TABLE}` (
                    id BIGINT PRIMARY KEY AUTO_INCREMENT,
                    phone VARCHAR(32) NOT NULL UNIQUE,
                    upstream_url TEXT,
                    remark VARCHAR(255),
                    assigned_tinyint TINYINT(1) DEFAULT 0,
                    assigned_type VARCHAR(32),
                    assigned_ref VARCHAR(64),
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
                    latest_sms_code VARCHAR(64),
                    latest_sms_fetched_at TIMESTAMP NULL DEFAULT NULL,
                    redeemed_at TIMESTAMP NULL DEFAULT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            cursor.execute(
                f"""
                CREATE TABLE IF NOT EXISTS `{MAILBOX_LIBRARY_TABLE}` (
                    id BIGINT PRIMARY KEY AUTO_INCREMENT,
                    email VARCHAR(255) NOT NULL UNIQUE,
                    email_password TEXT,
                    client_id VARCHAR(255),
                    refresh_token TEXT NOT NULL,
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
            if "assigned_type" not in sms_phone_columns:
                cursor.execute(f"ALTER TABLE `{SMS_PHONE_TABLE}` ADD COLUMN assigned_type VARCHAR(32)")
            if "assigned_ref" not in sms_phone_columns:
                cursor.execute(f"ALTER TABLE `{SMS_PHONE_TABLE}` ADD COLUMN assigned_ref VARCHAR(64)")
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
            if "latest_sms_code" not in sms_cdk_columns:
                cursor.execute(f"ALTER TABLE `{SMS_CDK_TABLE}` ADD COLUMN latest_sms_code VARCHAR(64)")
            if "latest_sms_fetched_at" not in sms_cdk_columns:
                cursor.execute(f"ALTER TABLE `{SMS_CDK_TABLE}` ADD COLUMN latest_sms_fetched_at TIMESTAMP NULL DEFAULT NULL")
            if "redeemed_at" not in sms_cdk_columns:
                cursor.execute(f"ALTER TABLE `{SMS_CDK_TABLE}` ADD COLUMN redeemed_at TIMESTAMP NULL DEFAULT NULL")
            if "created_at" not in sms_cdk_columns:
                cursor.execute(f"ALTER TABLE `{SMS_CDK_TABLE}` ADD COLUMN created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP")

            cursor.execute(f"SHOW COLUMNS FROM `{MAILBOX_LIBRARY_TABLE}`")
            mailbox_columns = {row["Field"] for row in cursor.fetchall()}
            if "email_password" not in mailbox_columns:
                cursor.execute(f"ALTER TABLE `{MAILBOX_LIBRARY_TABLE}` ADD COLUMN email_password TEXT")
            if "client_id" not in mailbox_columns:
                cursor.execute(f"ALTER TABLE `{MAILBOX_LIBRARY_TABLE}` ADD COLUMN client_id VARCHAR(255)")
            if "refresh_token" not in mailbox_columns:
                cursor.execute(f"ALTER TABLE `{MAILBOX_LIBRARY_TABLE}` ADD COLUMN refresh_token TEXT NOT NULL")
            if "created_at" not in mailbox_columns:
                cursor.execute(f"ALTER TABLE `{MAILBOX_LIBRARY_TABLE}` ADD COLUMN created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP")
        conn.commit()
    finally:
        conn.close()


def safe_schema_bootstrap():
    try:
        init_db()
    except Exception as exc:
        print(f"[startup] init_db failed: {exc}")

    # Best-effort lightweight migrations against an existing DB connection.
    try:
        conn = get_db_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute(f"SHOW COLUMNS FROM `{TABLE_NAME}`")
                account_columns = {row["Field"] for row in cursor.fetchall()}
                if "phone" not in account_columns:
                    cursor.execute(f"ALTER TABLE `{TABLE_NAME}` ADD COLUMN phone VARCHAR(32)")

                cursor.execute(f"SHOW COLUMNS FROM `{SMS_PHONE_TABLE}`")
                sms_phone_columns = {row["Field"] for row in cursor.fetchall()}
                if "assigned_type" not in sms_phone_columns:
                    cursor.execute(f"ALTER TABLE `{SMS_PHONE_TABLE}` ADD COLUMN assigned_type VARCHAR(32)")
                if "assigned_ref" not in sms_phone_columns:
                    cursor.execute(f"ALTER TABLE `{SMS_PHONE_TABLE}` ADD COLUMN assigned_ref VARCHAR(64)")

            conn.commit()
        finally:
            conn.close()
    except Exception as exc:
        print(f"[startup] safe_schema_bootstrap failed: {exc}")


def generate_random_code(prefix="GPT", length=12):
    alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
    chunks = []
    chunk_sizes = [4, 4, max(4, length - 8)]
    for size in chunk_sizes:
        chunks.append("".join(random.choice(alphabet) for _ in range(size)))
    return f"{prefix}-{chunks[0]}-{chunks[1]}-{chunks[2]}"


def generate_unique_cdk(prefix="GPT", length=12, table=TABLE_NAME, field="cdk", cursor=None):
    if cursor is not None:
        for _ in range(50):
            code = generate_random_code(prefix=prefix, length=length)
            cursor.execute(f"SELECT 1 FROM `{table}` WHERE `{field}`=%s", (code,))
            if not cursor.fetchone():
                return code
        raise RuntimeError("生成随机 CDK 失败，请重试")

    conn = get_db_connection()
    try:
        with conn.cursor() as cursor_local:
            for _ in range(50):
                code = generate_random_code(prefix=prefix, length=length)
                cursor_local.execute(f"SELECT 1 FROM `{table}` WHERE `{field}`=%s", (code,))
                if not cursor_local.fetchone():
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


# 内存 token 缓存：避免每次请求都重新换取 access_token（有效期通常 3600s）
_token_cache: dict = {}


def _token_cache_key(refresh_token_val: str, client_id_val: str) -> str:
    """对 refresh_token 做 sha256 摘要，避免把长 token 当 dict key。"""
    raw = f"{client_id_val}:{refresh_token_val}"
    return hashlib.sha256(raw.encode()).hexdigest()


def fetch_graph_access_token(refresh_token_val, client_id_val):
    """获取 Graph API access_token，优先使用内存缓存，过期前 60s 主动刷新。"""
    cache_key = _token_cache_key(refresh_token_val, client_id_val)
    cached = _token_cache.get(cache_key)
    if cached and time.time() < cached["_expires_at"] - 60:
        return cached
    result = fetch_access_token(
        refresh_token_val, client_id_val, "https://graph.microsoft.com/.default"
    )
    expires_in = int(result.get("expires_in") or 3600)
    _token_cache[cache_key] = {**result, "_expires_at": time.time() + expires_in}
    return _token_cache[cache_key]


def graph_headers(access_token):
    return {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}


def extract_verification_code(text):
    patterns = [
        # Match 'verification code is: 324177' or similar
        r"(?:verification code|security code|one-time code|one-time password|otp|验证码|校验码|安全代码)\s+(?:is|是)[:：\s-]*([A-Za-z0-9]{4,10})",
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
        "date": date_obj.isoformat(),
        "code": extract_verification_code(f"{subject_str}\n{body_preview}"),
        "date_obj": date_obj,
    }


def filter_recent_items(items, minutes=RECENT_MINUTES):
    cutoff = datetime.now(timezone.utc) - timedelta(minutes=minutes)
    return [item for item in items if item.get("date_obj") and item["date_obj"] >= cutoff]


def fetch_all_emails(email_address, refresh_token, client_id, max_per_mailbox=30, apply_recent_limit=True):
    graph_result = fetch_graph_access_token(refresh_token, client_id or DEFAULT_CLIENT_ID)
    access_token = graph_result.get("access_token")
    scope = graph_result.get("scope", "")
    if "https://graph.microsoft.com/Mail.Read" not in scope:
        raise RuntimeError(f"当前 token 不支持 Graph Mail.Read，返回 scope: {scope or '(empty)'}")

    raw_messages = fetch_graph_messages(access_token, max_items=max_per_mailbox)
    found_items = [parse_graph_message(message) for message in raw_messages]
    found_items.sort(key=lambda x: x["date_obj"], reverse=True)
    if apply_recent_limit:
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
                SELECT cdk, email, phone, email_password, client_id, gpt_password, refresh_token, redeemed
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
            cursor.execute(
                f"UPDATE `{TABLE_NAME}` SET redeemed = 1, redeemed_at = NOW() WHERE cdk = %s",
                (cdk,),
            )
        conn.commit()
    finally:
        conn.close()


def mark_bound_phone_redeemed(phone):
    if not phone:
        return
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                UPDATE `{SMS_CDK_TABLE}`
                SET status = 'redeemed', redeemed_at = NOW()
                WHERE phone = %s
                ORDER BY id DESC
                LIMIT 1
                """,
                (phone,),
            )
        conn.commit()
    finally:
        conn.close()


def mark_account_sold(cdk):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(f"UPDATE `{TABLE_NAME}` SET sold = 1 WHERE cdk = %s", (cdk,))
            deleted = cursor.rowcount
        conn.commit()
        return deleted
    finally:
        conn.close()


def list_accounts():
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT cdk, email, phone, gpt_password, redeemed, sold, redeemed_at, created_at
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


def assign_phone_to_account(cdk):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(f"SELECT cdk, phone FROM `{TABLE_NAME}` WHERE cdk = %s", (cdk,))
            account = cursor.fetchone()
            if not account:
                raise ValueError("未找到对应账号")
            if account.get("phone"):
                return {"cdk": cdk, "phone": account["phone"], "status": "existing"}

            cursor.execute(
                f"""
                SELECT phone
                FROM `{SMS_PHONE_TABLE}`
                WHERE (assigned_type IS NULL OR assigned_type = '')
                ORDER BY id ASC
                LIMIT 1
                """
            )
            phone_row = cursor.fetchone()
            if not phone_row:
                raise ValueError("没有可分配的手机号")

            phone = phone_row["phone"]
            cursor.execute(f"UPDATE `{TABLE_NAME}` SET phone = %s WHERE cdk = %s", (phone, cdk))
            cursor.execute(
                f"""
                UPDATE `{SMS_PHONE_TABLE}`
                SET assigned_tinyint = 1, assigned_type = 'account', assigned_ref = %s
                WHERE phone = %s
                """,
                (cdk, phone),
            )
        conn.commit()
        return {"cdk": cdk, "phone": phone, "status": "assigned"}
    finally:
        conn.close()


def unbind_phone_from_account(cdk):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(f"SELECT phone FROM `{TABLE_NAME}` WHERE cdk = %s", (cdk,))
            account = cursor.fetchone()
            if not account:
                raise ValueError("未找到对应账号")
            phone = account.get("phone")
            if not phone:
                return {"cdk": cdk, "phone": "", "status": "empty"}

            cursor.execute(f"UPDATE `{TABLE_NAME}` SET phone = NULL WHERE cdk = %s", (cdk,))
            cursor.execute(
                f"""
                UPDATE `{SMS_PHONE_TABLE}`
                SET assigned_tinyint = 0, assigned_type = NULL, assigned_ref = NULL
                WHERE phone = %s AND assigned_type = 'account' AND assigned_ref = %s
                """,
                (phone, cdk),
            )
        conn.commit()
        return {"cdk": cdk, "phone": phone, "status": "released"}
    finally:
        conn.close()


def import_account_lines(raw_lines):
    rows = []
    for line in raw_lines.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = [part.strip() for part in line.split("----")]
        if len(parts) != 4:
            raise ValueError("格式错误，必须为 邮箱----邮箱密码/授权码----client_id----refresh_token")
        rows.append(parts)

    if not rows:
        raise ValueError("没有可导入的邮箱账号")

    conn = get_db_connection()
    inserted = 0
    updated = 0
    results = []
    try:
        with conn.cursor() as cursor:
            for email, email_password, client_id, refresh_token in rows:
                cursor.execute(f"SELECT cdk FROM `{TABLE_NAME}` WHERE lower(email)=lower(%s)", (email,))
                existing = cursor.fetchone()
                if existing:
                    updated += 1
                    results.append({"email": email, "status": "skipped", "cdk": existing["cdk"]})
                    continue

                cdk = generate_unique_cdk(prefix="GPT", length=12, table=TABLE_NAME, field="cdk", cursor=cursor)
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
                inserted += 1
                results.append({"email": email, "status": "inserted", "cdk": cdk})
        conn.commit()
        return {
            "total": len(rows),
            "inserted": inserted,
            "updated": updated,
            "items": results,
        }
    finally:
        conn.close()


def add_mailbox_to_account_inventory(mailbox_id):
    mailbox = get_mailbox_library_item(mailbox_id)
    if not mailbox:
        raise ValueError("未找到对应邮箱")

    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(f"SELECT cdk FROM `{TABLE_NAME}` WHERE lower(email)=lower(%s)", (mailbox["email"],))
            existing = cursor.fetchone()
            if existing:
                return {"status": "exists", "cdk": existing["cdk"], "email": mailbox["email"]}

            cdk = generate_unique_cdk(prefix="GPT", length=12, table=TABLE_NAME, field="cdk", cursor=cursor)
            cursor.execute(
                f"""
                INSERT INTO `{TABLE_NAME}` (cdk, email, email_password, client_id, gpt_password, refresh_token)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (
                    cdk,
                    mailbox["email"],
                    mailbox.get("email_password") or "",
                    mailbox.get("client_id") or DEFAULT_CLIENT_ID,
                    os.getenv("DEFAULT_GPT_PASSWORD", "Jijie@123456"),
                    mailbox["refresh_token"],
                ),
            )
        conn.commit()
        return {"status": "inserted", "cdk": cdk, "email": mailbox["email"]}
    finally:
        conn.close()


def import_sms_lines(raw_lines):
    rows = []
    for line in raw_lines.splitlines():
        line = line.strip()
        if not line:
            continue
        if "----" in line:
            parts = [part.strip() for part in line.split("----")]
        elif "|" in line:
            parts = [part.strip() for part in line.split("|", 2)]
        else:
            parts = [part.strip() for part in line.split("----")]
        
        if len(parts) < 2:
            raise ValueError("手机号导入格式错误，必须为 手机号----上游链接 或 手机号|上游链接")
        phone = parts[0]
        upstream_url = parts[1]
        remark = parts[2] if len(parts) > 2 else ""
        rows.append((phone, upstream_url, remark))

    if not rows:
        raise ValueError("没有可导入的手机号")

    conn = get_db_connection()
    inserted = 0
    updated = 0
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
                else:
                    cursor.execute(
                        f"INSERT INTO `{SMS_PHONE_TABLE}` (phone, upstream_url, remark) VALUES (%s, %s, %s)",
                        (phone, upstream_url, remark),
                    )
                    inserted += 1
        conn.commit()
        return {
            "total": len(rows),
            "inserted": inserted,
            "updated": updated,
            "generated": 0,
            "codes": [],
        }
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def import_mailbox_library_lines(raw_lines):
    rows = []
    for line in raw_lines.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = [part.strip() for part in line.split("----")]
        if len(parts) != 4:
            raise ValueError("格式错误，必须为 邮箱----邮箱密码/授权码----client_id----refresh_token")
        rows.append(parts)

    if not rows:
        raise ValueError("没有可导入的邮箱")

    conn = get_db_connection()
    inserted = 0
    updated = 0
    try:
        with conn.cursor() as cursor:
            for email, email_password, client_id, refresh_token in rows:
                cursor.execute(f"SELECT id FROM `{MAILBOX_LIBRARY_TABLE}` WHERE lower(email)=lower(%s)", (email,))
                existing = cursor.fetchone()
                if existing:
                    cursor.execute(
                        f"""
                        UPDATE `{MAILBOX_LIBRARY_TABLE}`
                        SET email_password=%s, client_id=%s, refresh_token=%s
                        WHERE id=%s
                        """,
                        (email_password, client_id, refresh_token, existing["id"]),
                    )
                    updated += 1
                    continue
                cursor.execute(
                    f"""
                    INSERT INTO `{MAILBOX_LIBRARY_TABLE}` (email, email_password, client_id, refresh_token)
                    VALUES (%s, %s, %s, %s)
                    """,
                    (email, email_password, client_id, refresh_token),
                )
                inserted += 1
        conn.commit()
        return {"total": len(rows), "inserted": inserted, "updated": updated}
    finally:
        conn.close()


def list_mailbox_library():
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT
                    m.id,
                    m.email,
                    m.email_password,
                    m.client_id,
                    m.refresh_token,
                    m.created_at,
                    a.cdk AS stocked_cdk
                FROM `{MAILBOX_LIBRARY_TABLE}` m
                LEFT JOIN `{TABLE_NAME}` a ON lower(a.email) = lower(m.email)
                ORDER BY (a.cdk IS NOT NULL) ASC, m.id DESC
                """
            )
            return cursor.fetchall()
    finally:
        conn.close()


def get_mailbox_library_item(mailbox_id):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT id, email, email_password, client_id, refresh_token, created_at
                FROM `{MAILBOX_LIBRARY_TABLE}`
                WHERE id=%s
                """,
                (mailbox_id,),
            )
            return cursor.fetchone()
    finally:
        conn.close()


def delete_mailbox_library_item(mailbox_id):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(f"DELETE FROM `{MAILBOX_LIBRARY_TABLE}` WHERE id=%s", (mailbox_id,))
            deleted = cursor.rowcount
        conn.commit()
        return deleted
    finally:
        conn.close()


def fetch_mailbox_library_emails(mailbox_id):
    mailbox = get_mailbox_library_item(mailbox_id)
    if not mailbox:
        raise ValueError("未找到对应邮箱")

    mails = fetch_all_emails(
        mailbox["email"],
        mailbox["refresh_token"],
        mailbox.get("client_id") or DEFAULT_CLIENT_ID,
        max_per_mailbox=20,
        apply_recent_limit=False,
    )
    return {
        "mailbox": {
            "id": mailbox["id"],
            "email": mailbox["email"],
            "created_at": mailbox["created_at"],
        },
        "items": mails,
        "count": len(mails),
    }


def generate_sms_cdks(count, batch_name, prefix, length):
    conn = get_db_connection()
    generated = []
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT phone
                FROM `{SMS_PHONE_TABLE}`
                WHERE (assigned_type IS NULL OR assigned_type = '')
                ORDER BY id ASC
                LIMIT %s
                """,
                (count,),
            )
            phones = cursor.fetchall()
            if not phones:
                raise ValueError("没有可分配的手机号")

            for row in phones:
                phone = row["phone"]
                code = generate_unique_cdk(prefix=prefix or "SMS", length=max(length, 8), table=SMS_CDK_TABLE, field="code", cursor=cursor)
                cursor.execute(
                    f"""
                    INSERT INTO `{SMS_CDK_TABLE}` (code, phone, status, batch_name)
                    VALUES (%s, %s, 'unused', %s)
                    """,
                    (code, phone, batch_name),
                )
                cursor.execute(
                    f"""
                    UPDATE `{SMS_PHONE_TABLE}`
                    SET assigned_tinyint = 1, assigned_type = 'sms_cdk', assigned_ref = %s
                    WHERE phone=%s
                    """,
                    (code, phone),
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
            cursor.execute(f"SELECT COUNT(*) AS count FROM `{SMS_PHONE_TABLE}` WHERE assigned_type IS NULL OR assigned_type = ''")
            available = cursor.fetchone()["count"]
            cursor.execute(f"SELECT COUNT(*) AS count FROM `{SMS_PHONE_TABLE}` WHERE assigned_type IS NOT NULL AND assigned_type <> ''")
            assigned = cursor.fetchone()["count"]
            cursor.execute(f"SELECT COUNT(*) AS count FROM `{SMS_CDK_TABLE}` WHERE status = 'unused'")
            unused = cursor.fetchone()["count"]
            cursor.execute(f"SELECT COUNT(*) AS count FROM `{SMS_CDK_TABLE}` WHERE status = 'redeemed'")
            redeemed = cursor.fetchone()["count"]
            cursor.execute(
                f"""
                SELECT code, status, phone, latest_sms_code, latest_sms_fetched_at, redeemed_at
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
                        f"""
                        UPDATE `{SMS_PHONE_TABLE}`
                        SET assigned_tinyint = 0, assigned_type = NULL, assigned_ref = NULL
                        WHERE phone = %s AND assigned_type = 'sms_cdk'
                        """,
                        (phone,),
                    )
            deleted = cursor.rowcount
        conn.commit()
        return deleted
    finally:
        conn.close()


def refresh_sms_cdk(code):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT id, phone, status, batch_name, latest_sms_code, latest_sms_fetched_at, redeemed_at
                FROM `{SMS_CDK_TABLE}`
                WHERE code = %s
                """,
                (code,),
            )
            row = cursor.fetchone()
            if not row:
                raise ValueError("未找到对应 SMS CDK")

            new_code = generate_unique_cdk(prefix="SMS", length=10, table=SMS_CDK_TABLE, field="code", cursor=cursor)
            cursor.execute(
                f"""
                UPDATE `{SMS_CDK_TABLE}`
                SET code = %s,
                    status = 'unused',
                    latest_sms_code = NULL,
                    latest_sms_fetched_at = NULL,
                    redeemed_at = NULL
                WHERE id = %s
                """,
                (new_code, row["id"]),
            )
        conn.commit()
        return {"old_code": code, "new_code": new_code}
    finally:
        conn.close()


def list_sms_phones():
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT phone, upstream_url, remark, assigned_tinyint, assigned_type, assigned_ref, created_at
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
    code = extract_verification_code(raw_text)

    if code:
        conn = get_db_connection()
        try:
            with conn.cursor() as cursor:
                cursor.execute(
                    f"""
                    UPDATE `{SMS_CDK_TABLE}`
                    SET latest_sms_code = %s, latest_sms_fetched_at = NOW()
                    WHERE phone = %s
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (code, phone),
                )
            conn.commit()
        finally:
            conn.close()

    return {
        "code": code,
        "raw": raw_text,
        "source": upstream_url,
    }


def get_cached_sms_message(phone):
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT latest_sms_code, latest_sms_fetched_at
                FROM `{SMS_CDK_TABLE}`
                WHERE phone = %s
                ORDER BY id DESC
                LIMIT 1
                """,
                (phone,),
            )
            row = cursor.fetchone()
    finally:
        conn.close()

    if not row or not row.get("latest_sms_code"):
        return None

    return {
        "code": row["latest_sms_code"],
        "raw": "",
        "source": "cache",
        "fetched_at": row.get("latest_sms_fetched_at"),
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
    phone = request.args.get("phone", "").strip()
    code = request.args.get("code", "").strip()
    if not phone and not code:
        return "参数错误：手机号或 CDK 不能为空", 400

    return render_template("sms_pickup.html", phone=phone, code=code)


@app.route("/sms-pickup/raw")
def sms_pickup_raw():
    phone = request.args.get("phone", "").strip()
    code = request.args.get("code", "").strip()
    if not phone and not code:
        return "参数错误：手机号或 CDK 不能为空", 400

    conn = get_db_connection()
    row = None
    try:
        with conn.cursor() as cursor:
            if code:
                cursor.execute(
                    f"""
                    SELECT c.code, c.phone, c.status, c.redeemed_at, p.remark
                    FROM `{SMS_CDK_TABLE}` c
                    LEFT JOIN `{SMS_PHONE_TABLE}` p ON c.phone = p.phone
                    WHERE c.code = %s
                    """,
                    (code,),
                )
                row = cursor.fetchone()
            else:
                cursor.execute(
                    f"""
                    SELECT c.code, c.phone, c.status, c.redeemed_at, p.remark
                    FROM `{SMS_CDK_TABLE}` c
                    LEFT JOIN `{SMS_PHONE_TABLE}` p ON c.phone = p.phone
                    WHERE c.phone = %s
                    ORDER BY c.id DESC LIMIT 1
                    """,
                    (phone,),
                )
                row = cursor.fetchone()
                if not row:
                    cursor.execute(
                        f"SELECT phone, remark FROM `{SMS_PHONE_TABLE}` WHERE phone = %s",
                        (phone,),
                    )
                    phone_row = cursor.fetchone()
                    if phone_row:
                        row = {
                            "phone": phone_row["phone"],
                            "code": "-",
                            "status": "unused",
                            "redeemed_at": None,
                            "remark": phone_row["remark"],
                        }
    finally:
        conn.close()

    if not row:
        return "未找到对应记录", 404

    try:
        sms_payload = fetch_sms_message(row["phone"])
        raw_text = (sms_payload.get("raw") or "").strip()
    except Exception:
        raw_text = ""

    return raw_text or "没有获取到验证码"


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
        return jsonify({"code": "", "raw": "", "source": "", "error": ""})


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
    mark_bound_phone_redeemed(account.get("phone"))
    pickup_url = ""
    if account.get("phone"):
        _, pickup_url = build_sms_links(account["phone"])
    return jsonify(
        {
            "cdk": account["cdk"],
            "email": account["email"],
            "phone": account.get("phone") or "",
            "pickupUrl": pickup_url,
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
            max_per_mailbox=10,
        )
    except Exception as exc:
        return jsonify({"error": f"获取邮件失败: {exc}"}), 500

    verification_mails = [item for item in mails if is_verification_email(item)][:1]
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


@app.route("/api/account-with-mails", methods=["POST"])
def account_with_mails():
    """合并接口：一次请求完成 CDK 兑换 + 邮件获取，减少前端网络往返。"""
    data = request.get_json(silent=True) or {}
    cdk = (data.get("cdk") or "").strip()
    if not cdk:
        return jsonify({"error": "CDK 不能为空"}), 400

    account = get_account_by_cdk(cdk)
    if not account:
        return jsonify({"error": "未找到对应 CDK"}), 404

    mark_account_redeemed(cdk)

    try:
        mails = fetch_all_emails(
            account["email"],
            account["refresh_token"],
            account.get("client_id") or DEFAULT_CLIENT_ID,
            max_per_mailbox=10,
        )
    except Exception:
        mails = []

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
    return jsonify({
        "cdk": account["cdk"],
        "email": account["email"],
        "gpt_password": account["gpt_password"],
        "mails": {"items": slim_items, "count": len(slim_items)},
    })


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
            max_per_mailbox=10,
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


@app.route("/api/admin/accounts/<cdk>/sell", methods=["POST"])
def admin_sell_account(cdk):
    guard = require_admin_api()
    if guard:
        return guard
    try:
        deleted = mark_account_sold(cdk)
    except Exception as exc:
        return jsonify({"error": f"标记出售失败: {exc}"}), 500
    if not deleted:
        return jsonify({"error": "未找到对应账号"}), 404
    return jsonify({"sold": True, "cdk": cdk})


@app.route("/api/admin/accounts/<cdk>/assign-phone", methods=["POST"])
def admin_assign_account_phone(cdk):
    guard = require_admin_api()
    if guard:
        return guard
    try:
        result = assign_phone_to_account(cdk)
        return jsonify(result)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"分配手机号失败: {exc}"}), 500


@app.route("/api/admin/accounts/<cdk>/unbind-phone", methods=["POST"])
def admin_unbind_account_phone(cdk):
    guard = require_admin_api()
    if guard:
        return guard
    try:
        result = unbind_phone_from_account(cdk)
        return jsonify(result)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"解绑手机号失败: {exc}"}), 500


@app.route("/api/admin/accounts/bulk-delete", methods=["POST"])
def admin_bulk_delete_accounts():
    guard = require_admin_api()
    if guard:
        return guard
    data = request.get_json(silent=True) or {}
    cdks = data.get("cdks") or []
    if not isinstance(cdks, list) or not cdks:
        return jsonify({"error": "请先选择要删除的账号"}), 400

    deleted = 0
    skipped = []
    for cdk in cdks:
        cdk = (cdk or "").strip()
        if not cdk:
            continue
        try:
            count = delete_account(cdk)
            if count:
                deleted += 1
            else:
                skipped.append(cdk)
        except Exception:
            skipped.append(cdk)

    return jsonify({"deleted": deleted, "skipped": skipped})


@app.route("/api/admin/import", methods=["POST"])
def admin_import():
    guard = require_admin_api()
    if guard:
        return guard
    data = request.get_json(silent=True) or {}
    raw_lines = (data.get("lines") or data.get("line") or "").strip()
    if not raw_lines:
        return jsonify({"error": "导入内容不能为空"}), 400
    try:
        result = import_account_lines(raw_lines)
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


@app.route("/api/admin/sms/cdks/<code>/refresh", methods=["POST"])
def admin_refresh_sms_cdk(code):
    guard = require_admin_api()
    if guard:
        return guard
    try:
        result = refresh_sms_cdk(code)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 404
    except Exception as exc:
        return jsonify({"error": f"刷新失败: {exc}"}), 500
    return jsonify(result)


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


@app.route("/api/admin/mailboxes", methods=["GET"])
def admin_mailboxes():
    guard = require_admin_api()
    if guard:
        return guard
    try:
        return jsonify({"items": list_mailbox_library()})
    except Exception as exc:
        return jsonify({"error": f"读取邮箱库失败: {exc}"}), 500


@app.route("/api/admin/mailboxes/import", methods=["POST"])
def admin_mailboxes_import():
    guard = require_admin_api()
    if guard:
        return guard
    data = request.get_json(silent=True) or {}
    lines = (data.get("lines") or "").strip()
    if not lines:
        return jsonify({"error": "导入内容不能为空"}), 400
    try:
        result = import_mailbox_library_lines(lines)
        return jsonify(result)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"导入邮箱库失败: {exc}"}), 500


@app.route("/api/admin/mailboxes/<int:mailbox_id>/mails", methods=["GET"])
def admin_mailbox_mails(mailbox_id):
    guard = require_admin_api()
    if guard:
        return guard
    try:
        return jsonify(fetch_mailbox_library_emails(mailbox_id))
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 404
    except Exception as exc:
        return jsonify({"error": f"获取邮箱邮件失败: {exc}"}), 500


@app.route("/api/admin/mailboxes/<int:mailbox_id>", methods=["DELETE"])
def admin_delete_mailbox(mailbox_id):
    guard = require_admin_api()
    if guard:
        return guard
    try:
        deleted = delete_mailbox_library_item(mailbox_id)
    except Exception as exc:
        return jsonify({"error": f"删除邮箱失败: {exc}"}), 500
    if not deleted:
        return jsonify({"error": "未找到对应邮箱"}), 404
    return jsonify({"deleted": True, "id": mailbox_id})


@app.route("/api/admin/mailboxes/<int:mailbox_id>/add-to-stock", methods=["POST"])
def admin_add_mailbox_to_stock(mailbox_id):
    guard = require_admin_api()
    if guard:
        return guard
    try:
        result = add_mailbox_to_account_inventory(mailbox_id)
        return jsonify(result)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 404
    except Exception as exc:
        return jsonify({"error": f"加入库存失败: {exc}"}), 500


if __name__ == "__main__":
    load_env_file()
    safe_schema_bootstrap()
    app.run(host="0.0.0.0", port=APP_PORT, debug=True)
