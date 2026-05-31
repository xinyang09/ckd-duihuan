import os
import sqlite3


DB_PATH = os.getenv("CDK_DB_PATH", os.path.join(os.getcwd(), "cdk.db"))
TABLE_NAME = os.getenv("CDK_TABLE_NAME", "cdk_accounts")

TEST_ROW = {
    "cdk": "TEST-CDK-1001",
    "email": "demo_outlook@example.com",
    "email_password": "DemoMailboxPass@123",
    "client_id": os.getenv("OUTLOOK_CLIENT_ID", ""),
    "gpt_password": "DemoGPT@123456",
    "refresh_token": "demo_refresh_token_replace_me",
}


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
                TEST_ROW["cdk"],
                TEST_ROW["email"],
                TEST_ROW["email_password"],
                TEST_ROW["client_id"],
                TEST_ROW["gpt_password"],
                TEST_ROW["refresh_token"],
            ),
        )
        conn.commit()
        print(f"测试数据已写入: {TEST_ROW['cdk']}")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
