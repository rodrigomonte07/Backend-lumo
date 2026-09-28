"""
Backend de geração de propostas — Rede Lumo.

Dois endpoints, exatamente como descritos no prompt do Lovable:

  POST /api/propostas/gerar   -> gera o .pptx a partir do template + dados do CRM
  POST /api/propostas/pdf     -> converte um .pptx já gerado em .pdf

Rodar localmente:
    pip install -r requirements.txt
    uvicorn app.main:app --reload --port 8000

O Lovable (ou qualquer front-end) chama esses endpoints por HTTP. Nenhuma
lógica de layout/design vive no front-end — tudo isso está resolvido aqui,
em cima do arquivo .pptx original do Canva.
"""

import json
import subprocess
import uuid
from pathlib import Path

from fastapi import FastAPI, Form, HTTPException, UploadFile
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
from pptx import Presentation

from import mail_merge

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = BASE_DIR / "template" / "modelo_rede_lumo.pptx"
FIELD_MAP_PATH = BASE_DIR / "template" / "field_map.json"
STORAGE_DIR = BASE_DIR / "storage"
PPTX_DIR = STORAGE_DIR / "pptx"
PDF_DIR = STORAGE_DIR / "pdf"
UPLOADS_DIR = STORAGE_DIR / "uploads"

for d in (PPTX_DIR, PDF_DIR, UPLOADS_DIR):
    d.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Rede Lumo — Geração de Propostas")

# Deixa os arquivos gerados acessíveis por URL direta (troque por S3/GCS em produção)
app.mount("/files/pptx", StaticFiles(directory=str(PPTX_DIR)), name="pptx-files")
app.mount("/files/pdf", StaticFiles(directory=str(PDF_DIR)), name="pdf-files")


@app.post("/api/propostas/gerar")
async def gerar_proposta(
    registro_id: str = Form(...),
    dados: str = Form(...),  # JSON string com os tokens (ver field_map.json)
    foto_fachada: UploadFile | None = None,
):
    try:
        data = json.loads(dados)
    except json.JSONDecodeError as e:
        raise HTTPException(400, f"'dados' não é um JSON válido: {e}")

    field_map = json.loads(FIELD_MAP_PATH.read_text(encoding="utf-8"))

    if not TEMPLATE_PATH.exists():
        raise HTTPException(500, "Template .pptx não encontrado no servidor.")

    prs = Presentation(str(TEMPLATE_PATH))
    report = {"applied": 0, "not_found": [], "missing_tokens": set(), "errors": [], "photo": None}
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


@app.post("/api/propostas/pdf")
async def gerar_pdf(pptx_id: str = Form(...)):
    """
    Converte um .pptx já gerado (por /gerar) em .pdf.

    ATENÇÃO — leia antes de usar em produção:
    Esta implementação usa LibreOffice headless (`soffice`), que é gratuito
    mas tem um bug de renderização CONFIRMADO neste template específico
    (ver slide "Valoração da Transação" — sobreposição de texto e corte de
    dígito). O mesmo arquivo abre perfeito no PowerPoint/Canva.
    Antes de ir para produção, troque a função `_convert_with_libreoffice`
    abaixo por um serviço de conversão mais fiel — ex: Aspose.Slides Cloud,
    Adobe PDF Services API, ou CloudConvert com engine que não seja
    LibreOffice puro. A interface da função (recebe caminho do .pptx,
    devolve caminho do .pdf) é a mesma — só troca a implementação interna.
    """
    src = PPTX_DIR / f"{pptx_id}.pptx"
    if not src.exists():
        raise HTTPException(404, f"pptx_id '{pptx_id}' não encontrado. Gere a proposta primeiro em /api/propostas/gerar.")

    pdf_path = _convert_with_libreoffice(src)

    return JSONResponse(
        {
            "status": "ok",
            "pdf_id": pptx_id,
            "pdf_url": f"/files/pdf/{pptx_id}.pdf",
            "aviso": "Convertido com LibreOffice — valide visualmente antes de produção (ver docstring do endpoint).",
        }
    )


def _convert_with_libreoffice(src_path: Path) -> Path:
    result = subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(PDF_DIR), str(src_path)],
        capture_output=True,
        text=True,
        timeout=90,
    )
    if result.returncode != 0:
        raise HTTPException(500, f"Falha ao converter para PDF: {result.stderr}")
    pdf_path = PDF_DIR / f"{src_path.stem}.pdf"
    if not pdf_path.exists():
        raise HTTPException(500, "Conversão terminou sem erro mas o PDF não foi encontrado.")
    return pdf_path


@app.get("/health")
def health():
    return {"status": "ok", "template_encontrado": TEMPLATE_PATH.exists()}
