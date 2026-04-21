import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

_ROOT = Path(__file__).resolve().parent.parent
_INSTANCE = _ROOT / "instance"


class Config:
    SECRET_KEY = os.environ.get("SECRET_KEY") or "dev-only-change-me"
    SQLALCHEMY_DATABASE_URI = os.environ.get("DATABASE_URL") or (
        f"sqlite:///{_INSTANCE / 'studysync.db'}"
    )
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    WTF_CSRF_ENABLED = True


class TestConfig(Config):
    TESTING = True
    WTF_CSRF_ENABLED = False
    SQLALCHEMY_DATABASE_URI = "sqlite:///:memory:"
