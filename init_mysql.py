import os

from app import init_db, load_env_file


def main():
    load_env_file()
    init_db()
    print("MySQL database and tables initialized.")


if __name__ == "__main__":
    main()
