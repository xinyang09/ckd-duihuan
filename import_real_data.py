import os
import sqlite3


DB_PATH = os.getenv("CDK_DB_PATH", os.path.join(os.getcwd(), "cdk.db"))
TABLE_NAME = os.getenv("CDK_TABLE_NAME", "cdk_accounts")

REAL_ROWS = [
    {
        "cdk": os.getenv("IMPORT_REAL_CDK", "REAL-CDK-1001"),
        "email": os.getenv("IMPORT_REAL_EMAIL", "replace_me@example.com"),
        "email_password": os.getenv("IMPORT_REAL_EMAIL_PASSWORD", "replace_me"),
        "client_id": os.getenv("IMPORT_REAL_CLIENT_ID", ""),
        "gpt_password": os.getenv("IMPORT_REAL_GPT_PASSWORD", "Jijie@123456"),
        "refresh_token": os.getenv("IMPORT_REAL_REFRESH_TOKEN", "replace_me"),
    }
]


def main():
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute(
            f"""
            CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cdk TEXT NOT NULL UNIQUE,
                email TEXT NOT NULL,
                email_password TEXT DEFAULT '',
                client_id TEXT DEFAULT '',
                gpt_password TEXT NOT NULL,
                refresh_token TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        columns = {row[1] for row in conn.execute(f"PRAGMA table_info({TABLE_NAME})").fetchall()}
        if "email_password" not in columns:
            conn.execute(f"ALTER TABLE {TABLE_NAME} ADD COLUMN email_password TEXT DEFAULT ''")
        if "client_id" not in columns:
            conn.execute(f"ALTER TABLE {TABLE_NAME} ADD COLUMN client_id TEXT DEFAULT ''")
        for row in REAL_ROWS:
            conn.execute(
                f"""
                INSERT INTO {TABLE_NAME} (cdk, email, email_password, client_id, gpt_password, refresh_token)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(cdk) DO UPDATE SET
                    email = excluded.email,
                    email_password = excluded.email_password,
                    client_id = excluded.client_id,
                    gpt_password = excluded.gpt_password,
                    refresh_token = excluded.refresh_token
                """,
                (
                    row["cdk"],
                    row["email"],
                    row["email_password"],
                    row["client_id"],
                    row["gpt_password"],
                    row["refresh_token"],
                ),
            )
        conn.commit()
        print(f"已写入 {len(REAL_ROWS)} 条真实数据")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
