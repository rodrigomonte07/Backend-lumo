"""Atalho de compatibilidade: a implementação vive em app/mail_merge.py."""
from app.mail_merge import *  # noqa: F401,F403
from app.mail_merge import main

if __name__ == "__main__":
    main()
