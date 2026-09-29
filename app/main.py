import json
import logging
import subprocess
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
from pptx import Presentation

from . import mail_merge
from .config import (
    CORS_ORIGINS,
    DEBUG,
    ENVIRONMENT,
    FIELD_MAP_PATH,
    LIBREOFFICE_PATH,
    LOG_LEVEL,
    PDF_TIMEOUT,
    STORAGE_DIR,
    TEMPLATE_PATH,
)
from .storage import storage_service

logging.basicConfig(
    level=getattr(logging, LOG_LEVEL.upper(), logging.INFO),
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger("rede_lumo")

storage_service.ensure_dirs()

app = FastAPI(
    title="Rede Lumo — Geração de Propostas",
    description="Backend para geração e conversão de propostas em PDF",
    version="1.0.0",
    debug=DEBUG,
)

if CORS_ORIGINS:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=CORS_ORIGINS,
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

if storage_service.backend == "local":
    app.mount("/files", StaticFiles(directory=str(STORAGE_DIR)), name="storage-files")


@app.get("/health")
def health():
    template_ok = TEMPLATE_PATH.exists()
    field_map_ok = FIELD_MAP_PATH.exists()
    payload = {
        "status": "ok" if template_ok and field_map_ok else "degraded",
        "environment": ENVIRONMENT,
        "template_encontrado": template_ok,
        "field_map_encontrado": field_map_ok,
        "storage_backend": storage_service.backend,
    }
    return payload, 200 if template_ok and field_map_ok else 503


@app.post("/api/propostas/gerar")
async def gerar_proposta(
    registro_id: str = Form(...),
    dados: str = Form(...),
    foto_fachada: UploadFile | None = None,
):
    if not registro_id or not registro_id.strip():
        raise HTTPException(400, "'registro_id' é obrigatório.")

    try:
        raw_data = json.loads(dados)
    except json.JSONDecodeError as exc:
        raise HTTPException(400, f"'dados' não é um JSON válido: {exc}") from exc

    if not isinstance(raw_data, dict):
        raise HTTPException(400, "'dados' deve ser um objeto JSON (dicionário).")

    if not FIELD_MAP_PATH.exists():
        raise HTTPException(500, "field_map.json não encontrado no servidor.")
    if not TEMPLATE_PATH.exists():
        raise HTTPException(500, "Template .pptx não encontrado no servidor.")

    field_map = json.loads(FIELD_MAP_PATH.read_text(encoding="utf-8"))
    prs = Presentation(str(TEMPLATE_PATH))

    report = {
        "applied": 0,
        "not_found": [],
        "missing_tokens": set(),
        "errors": [],
        "photo": None,
    }

    try:
        mail_merge.apply_field_map(prs, field_map, raw_data, report)
    except Exception as exc:  # pragma: no cover - defensive guard
        logger.exception("Erro ao aplicar field_map")
        raise HTTPException(500, f"Erro ao aplicar field_map: {exc}") from exc

    file_id = uuid.uuid4().hex
    temp_pptx = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
    temp_pptx.close()
    pptx_path = Path(temp_pptx.name)
    try:
        prs.save(str(pptx_path))

        if foto_fachada is not None:
            file_bytes = await foto_fachada.read()
            if not file_bytes:
                raise HTTPException(400, "Arquivo de foto recebido está vazio.")
            photo_extension = Path(foto_fachada.filename or "photo.jpg").suffix or ".jpg"
            with tempfile.NamedTemporaryFile(suffix=photo_extension, delete=False) as photo_tmp:
                photo_tmp.write(file_bytes)
                photo_path = Path(photo_tmp.name)
            try:
                mail_merge.replace_photo(str(pptx_path), str(photo_path), report)
            finally:
                if photo_path.exists():
                    photo_path.unlink(missing_ok=True)

        if not report["photo"]:
            report["photo"] = "nenhuma foto enviada — mantida a foto original do template"

        pptx_bytes = pptx_path.read_bytes()
        stored_rel = f"pptx/{file_id}.pptx"
        pptx_url = storage_service.save_bytes(stored_rel, pptx_bytes)

        status = "ok" if not report["missing_tokens"] and not report["errors"] else "ok_com_avisos"

        return JSONResponse(
            {
                "status": status,
                "pptx_id": file_id,
                "pptx_url": pptx_url,
                "registro_id": registro_id,
                "campos_aplicados": report["applied"],
                "avisos": {
                    "tokens_faltando": sorted(report["missing_tokens"]),
                    "regras_nao_encontradas": report["not_found"],
                    "erros": report["errors"],
                },
            }
        )
    finally:
        if pptx_path.exists():
            pptx_path.unlink(missing_ok=True)


@app.post("/api/propostas/pdf")
async def gerar_pdf(pptx_id: str = Form(...)):
    if not pptx_id or not pptx_id.strip():
        raise HTTPException(400, "'pptx_id' é obrigatório.")

    stored_key = f"pptx/{pptx_id}.pptx"
    if not storage_service.file_exists(stored_key):
        raise HTTPException(
            404,
            f"pptx_id '{pptx_id}' não encontrado. Gere a proposta primeiro em /api/propostas/gerar.",
        )

    src_bytes = storage_service.read_bytes(stored_key)
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp_src:
        tmp_src.write(src_bytes)
        src_path = Path(tmp_src.name)

    try:
        pdf_path = _convert_with_libreoffice(src_path)
        pdf_bytes = pdf_path.read_bytes()
        pdf_key = f"pdf/{pptx_id}.pdf"
        pdf_url = storage_service.save_bytes(pdf_key, pdf_bytes)
    finally:
        if src_path.exists():
            src_path.unlink(missing_ok=True)
        if (Path(tempfile.gettempdir()) / f"{pptx_id}.pdf").exists():
            (Path(tempfile.gettempdir()) / f"{pptx_id}.pdf").unlink(missing_ok=True)

    return JSONResponse(
        {
            "status": "ok",
            "pdf_id": pptx_id,
            "pdf_url": pdf_url,
            "aviso": "Convertido com LibreOffice — valide visualmente antes de produção.",
        }
    )


def _convert_with_libreoffice(src_path: Path) -> Path:
    try:
        result = subprocess.run(
            [
                LIBREOFFICE_PATH,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(Path(tempfile.gettempdir())),
                str(src_path),
            ],
            capture_output=True,
            text=True,
            timeout=PDF_TIMEOUT,
        )
    except FileNotFoundError as exc:
        raise HTTPException(
            500,
            f"LibreOffice não encontrado. Instale 'libreoffice' e configure LIBREOFFICE_PATH. Detalhe: {exc}",
        ) from exc
    except subprocess.TimeoutExpired as exc:
        raise HTTPException(500, f"Timeout ao converter para PDF (limite: {PDF_TIMEOUT}s)") from exc

    if result.returncode != 0:
        raise HTTPException(500, f"Falha ao converter para PDF: {result.stderr or 'sem mensagem'}")

    pdf_path = Path(tempfile.gettempdir()) / f"{src_path.stem}.pdf"
    if not pdf_path.exists():
        raise HTTPException(500, "Conversão terminou sem erro mas o PDF não foi encontrado.")

    return pdf_path


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=DEBUG)
