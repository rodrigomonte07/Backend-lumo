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

# Garantir que a pasta storage exista
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
