import json
import logging
import subprocess
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
    PDF_DIR,
    PDF_TIMEOUT,
    PPTX_DIR,
    STORAGE_DIR,
    TEMPLATE_PATH,
    UPLOADS_DIR,
)

logging.basicConfig(
    level=getattr(logging, LOG_LEVEL.upper(), logging.INFO),
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger("rede_lumo")

for directory in (PPTX_DIR, PDF_DIR, UPLOADS_DIR):
    directory.mkdir(parents=True, exist_ok=True)

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

app.mount("/files/pptx", StaticFiles(directory=str(PPTX_DIR)), name="pptx-files")
app.mount("/files/pdf", StaticFiles(directory=str(PDF_DIR)), name="pdf-files")


@app.get("/health")
def health():
    template_ok = TEMPLATE_PATH.exists()
    field_map_ok = FIELD_MAP_PATH.exists()
    if not template_ok or not field_map_ok:
        return {
            "status": "degraded",
            "environment": ENVIRONMENT,
            "template_encontrado": template_ok,
            "field_map_encontrado": field_map_ok,
        }, 503
    return {
        "status": "ok",
        "environment": ENVIRONMENT,
        "template_encontrado": template_ok,
        "field_map_encontrado": field_map_ok,
    }


@app.post("/api/propostas/gerar")
async def gerar_proposta(
    registro_id: str = Form(...),
    dados: str = Form(...),
    foto_fachada: UploadFile | None = None,
):
    try:
        try:
            data = json.loads(dados)
        except json.JSONDecodeError as exc:
            raise HTTPException(400, f"'dados' não é um JSON válido: {exc}")

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

        mail_merge.apply_field_map(prs, field_map, data, report)

        file_id = uuid.uuid4().hex
        out_path = PPTX_DIR / f"{file_id}.pptx"
        prs.save(str(out_path))

        if foto_fachada is not None:
            upload_path = UPLOADS_DIR / f"{file_id}_{foto_fachada.filename}"
            upload_path.write_bytes(await foto_fachada.read())
            mail_merge.replace_photo(str(out_path), str(upload_path), report)
        else:
            report["photo"] = "nenhuma foto enviada — mantida a foto original do template"

        status = "ok" if not report["missing_tokens"] and not report["errors"] else "ok_com_avisos"

        return JSONResponse(
            {
                "status": status,
                "pptx_id": file_id,
                "pptx_url": f"/files/pptx/{file_id}.pptx",
                "registro_id": registro_id,
                "campos_aplicados": report["applied"],
                "avisos": {
                    "tokens_faltando": sorted(report["missing_tokens"]),
                    "regras_nao_encontradas": report["not_found"],
                    "erros": report["errors"],
                },
            }
        )

    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("Erro interno em /api/propostas/gerar")
        raise HTTPException(500, f"Erro interno: {exc}") from exc


@app.post("/api/propostas/pdf")
async def gerar_pdf(pptx_id: str = Form(...)):
    src = PPTX_DIR / f"{pptx_id}.pptx"
    if not src.exists():
        raise HTTPException(
            404,
            f"pptx_id '{pptx_id}' não encontrado. Gere a proposta primeiro em /api/propostas/gerar.",
        )

    pdf_path = _convert_with_libreoffice(src)

    return JSONResponse(
        {
            "status": "ok",
            "pdf_id": pptx_id,
            "pdf_url": f"/files/pdf/{pptx_id}.pdf",
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
                str(PDF_DIR),
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

    pdf_path = PDF_DIR / f"{src_path.stem}.pdf"
    if not pdf_path.exists():
        raise HTTPException(500, "Conversão terminou sem erro mas o PDF não foi encontrado.")

    return pdf_path


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=DEBUG)
