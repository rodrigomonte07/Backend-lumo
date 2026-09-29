from __future__ import annotations

import os
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent


def _first_existing(*paths: Path) -> Path:
    """Retorna o primeiro caminho que existe, ou o primeiro da lista se nenhum existir."""
    for path in paths:
        if path.exists():
            return path
    return paths[0]


TEMPLATE_CANDIDATES = (
    BASE_DIR / "template" / "modelo_rede_lumo.pptx",
    BASE_DIR / "template" / "modelo_rede_lumo (1).pptx",
    BASE_DIR / "modelo_rede_lumo.pptx",
    BASE_DIR / "modelo_rede_lumo (1).pptx",
)

FIELD_MAP_CANDIDATES = (
    BASE_DIR / "template" / "field_map.json",
    BASE_DIR / "field_map.json",
)

STORAGE_DIR = BASE_DIR / "storage"
PPTX_DIR = STORAGE_DIR / "pptx"
PDF_DIR = STORAGE_DIR / "pdf"
UPLOADS_DIR = STORAGE_DIR / "uploads"

TEMPLATE_PATH = _first_existing(*TEMPLATE_CANDIDATES)
FIELD_MAP_PATH = _first_existing(*FIELD_MAP_CANDIDATES)

ENVIRONMENT = os.getenv("ENVIRONMENT", "development")
DEBUG = ENVIRONMENT == "development"
LOG_LEVEL = os.getenv("LOG_LEVEL", "DEBUG" if DEBUG else "INFO")
LIBREOFFICE_PATH = os.getenv("LIBREOFFICE_PATH", "soffice")
PDF_TIMEOUT = int(os.getenv("PDF_TIMEOUT", "90"))

CORS_ORIGINS = [
    origin.strip()
    for origin in os.getenv("CORS_ORIGINS", "*").split(",")
    if origin.strip()
]

AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID", "").strip()
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY", "").strip()
AWS_REGION = os.getenv("AWS_REGION", os.getenv("AWS_S3_REGION", "us-east-1")).strip()
AWS_S3_BUCKET = os.getenv("AWS_S3_BUCKET", "").strip()
AWS_S3_PUBLIC_URL = os.getenv("AWS_S3_PUBLIC_URL", "").strip()
STORAGE_BACKEND = os.getenv("STORAGE_BACKEND", "local").strip().lower()

if STORAGE_BACKEND == "s3" and not AWS_S3_BUCKET:
    STORAGE_BACKEND = "local"

if STORAGE_BACKEND not in {"local", "s3"}:
    STORAGE_BACKEND = "local"
