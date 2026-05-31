# CDK 邮箱查询 Web 端

## 功能

- 输入 CDK 查询数据库中的邮箱和 GPT 密码
- 点击按钮通过 Outlook OAuth + IMAP 获取最新验证码邮件
- 同时检查 `INBOX` 和 `Junk`

## 数据库

默认使用当前目录下的 `cdk.db`，表名默认 `cdk_accounts`。

字段结构：

- `cdk`
- `email`
- `gpt_password`
- `refresh_token`

可通过环境变量覆盖：

- `CDK_DB_PATH`
- `CDK_TABLE_NAME`
- `OUTLOOK_CLIENT_ID`
- `OUTLOOK_IMAP_SERVER`
- `OUTLOOK_IMAP_PORT`
- `OUTLOOK_TOKEN_URL`

## 启动

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python app.py
```

打开 `http://127.0.0.1:5090`

## Docker Compose 启动

项目支持直接通过 `docker compose` 启动，默认读取当前目录下的 `.env`。

```bash
docker compose up -d --build
```

查看日志：

```bash
docker compose logs -f
```

停止服务：

```bash
docker compose down
```
