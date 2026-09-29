from __future__ import annotations

from pathlib import Path

import boto3

from .config import (
    AWS_REGION,
    AWS_S3_BUCKET,
    AWS_S3_PUBLIC_URL,
    STORAGE_BACKEND,
    STORAGE_DIR,
)


class StorageService:
    def __init__(self) -> None:
        self.backend = STORAGE_BACKEND
        self.s3_client = None

        if self.backend == "s3" and AWS_S3_BUCKET:
            try:
                self.s3_client = boto3.client("s3", region_name=AWS_REGION)
            except Exception:
                self.backend = "local"

    def ensure_dirs(self) -> None:
        if self.backend != "local":
            return

        for directory in (STORAGE_DIR, STORAGE_DIR / "pptx", STORAGE_DIR / "pdf", STORAGE_DIR / "uploads"):
            directory.mkdir(parents=True, exist_ok=True)

    def save_bytes(self, relative_path: str, data: bytes) -> str:
        normalized = relative_path.lstrip("/")
        if self.backend == "s3" and AWS_S3_BUCKET:
            self.s3_client.put_object(Bucket=AWS_S3_BUCKET, Key=normalized, Body=data)
            return self.url_for(normalized)

        path = STORAGE_DIR / normalized
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        return f"/files/{normalized}"

    def read_bytes(self, relative_path: str) -> bytes:
        normalized = relative_path.lstrip("/")
        if self.backend == "s3" and AWS_S3_BUCKET:
            response = self.s3_client.get_object(Bucket=AWS_S3_BUCKET, Key=normalized)
            return response["Body"].read()
        return (STORAGE_DIR / normalized).read_bytes()

    def file_exists(self, relative_path: str) -> bool:
        normalized = relative_path.lstrip("/")
        if self.backend == "s3" and AWS_S3_BUCKET:
            try:
                self.s3_client.head_object(Bucket=AWS_S3_BUCKET, Key=normalized)
                return True
            except Exception:
                return False
        return (STORAGE_DIR / normalized).exists()

    def url_for(self, relative_path: str) -> str:
        normalized = relative_path.lstrip("/")
        if self.backend == "s3" and AWS_S3_BUCKET:
            if AWS_S3_PUBLIC_URL:
                return f"{AWS_S3_PUBLIC_URL.rstrip('/')}/{normalized}"
            return f"https://{AWS_S3_BUCKET}.s3.{AWS_REGION}.amazonaws.com/{normalized}"
        return f"/files/{normalized}"


storage_service = StorageService()
