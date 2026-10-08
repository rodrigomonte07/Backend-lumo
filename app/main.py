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
import os
import subprocess
import tempfile
import shutil
import uuid
from pathlib import Path

import httpx
from fastapi import FastAPI, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles

from . import mail_merge

BASE_DIR = Path(__file__).resolve().parent.parent

# Resolve template path with fallback: try template/ first, then root
def _find_first_existing(*paths):
    """Return the first existing path, or raise an error if none exist."""
    for p in paths:
        if p.exists():
            return p
    raise FileNotFoundError(
        f"None of the candidate paths exist:\n" + "\n".join(str(p) for p in paths)
    )

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

try:
    TEMPLATE_PATH = _find_first_existing(*TEMPLATE_CANDIDATES)
except FileNotFoundError as e:
    TEMPLATE_PATH = None
    TEMPLATE_ERROR = str(e)

try:
    FIELD_MAP_PATH = _find_first_existing(*FIELD_MAP_CANDIDATES)
except FileNotFoundError as e:
    FIELD_MAP_PATH = None
    FIELD_MAP_ERROR = str(e)

STORAGE_DIR = BASE_DIR / "storage"
PPTX_DIR = STORAGE_DIR / "pptx"
PDF_DIR = STORAGE_DIR / "pdf"
UPLOADS_DIR = STORAGE_DIR / "uploads"

for d in (PPTX_DIR, PDF_DIR, UPLOADS_DIR):
    d.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Rede Lumo — Geração de Propostas")

# URL pública do serviço (ex: https://meu-servico.onrender.com), sem barra no final.
# Usada para devolver links ABSOLUTOS de download ao front-end.
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "").rstrip("/")

# Origens permitidas a chamar a API pelo navegador (o app do Lovable).
# Ex: ALLOWED_ORIGINS="https://bloomexpansao.com.br,https://seu-projeto.lovable.app"
_origins = [o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "*").split(",") if o.strip()]
app.add_middleware(
    CORSMiddleware,
    allow_origins=_origins,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Deixa os arquivos gerados acessíveis por URL direta (troque por S3/GCS em produção)
app.mount("/files/pptx", StaticFiles(directory=str(PPTX_DIR)), name="pptx-files")
app.mount("/files/pdf", StaticFiles(directory=str(PDF_DIR)), name="pdf-files")


@app.post("/api/propostas/gerar")
async def gerar_proposta(
    registro_id: str = Form(...),
    dados: str = Form(...),  # JSON string com os tokens (ver field_map.json)
    foto_fachada: UploadFile | None = None,
):
    if TEMPLATE_PATH is None:
        raise HTTPException(500, f"Template .pptx não encontrado no servidor. {TEMPLATE_ERROR}")
    if FIELD_MAP_PATH is None:
        raise HTTPException(500, f"field_map.json não encontrado no servidor. {FIELD_MAP_ERROR}")

    try:
        data = json.loads(dados)
    except json.JSONDecodeError as e:
        raise HTTPException(400, f"'dados' não é um JSON válido: {e}")

    field_map = json.loads(FIELD_MAP_PATH.read_text(encoding="utf-8"))

    file_id = uuid.uuid4().hex
    out_path = PPTX_DIR / f"{file_id}.pptx"

    upload_path = None
    if foto_fachada is not None:
        upload_path = UPLOADS_DIR / f"{file_id}_{Path(foto_fachada.filename or 'foto').name}"
        upload_path.write_bytes(await foto_fachada.read())

    report = mail_merge.new_report()
    try:
        mail_merge.generate(TEMPLATE_PATH, field_map, data, out_path, str(upload_path) if upload_path else None, report)
    except Exception as e:  # foto corrompida, template inválido etc.
        raise HTTPException(500, f"Falha ao gerar a proposta: {e}")

    status = "ok" if not report["missing_tokens"] and not report["errors"] else "ok_com_avisos"

    return JSONResponse(
        {
            "status": status,
            "pptx_id": file_id,
            "pptx_url": f"{PUBLIC_BASE_URL}/files/pptx/{file_id}.pptx",
            "registro_id": registro_id,
            "campos_aplicados": report["applied"],
            "avisos": {
                "tokens_faltando": sorted(report["missing_tokens"]),
                "regras_nao_encontradas": report["not_found"],
                "erros": report["errors"],
                "foto": report["photo"],
                "valores_derivados": report.get("valores_derivados", []),
                "inconsistencias": report.get("inconsistencias", []),
                "ajustes_de_fonte": report.get("ajustes_de_fonte", []),
                "ajustes_de_largura": report.get("ajustes_de_largura", []),
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

    # Usa Aspose.Slides Cloud quando configurado (fiel ao PowerPoint, sem o bug
    # de fonte/sobreposição do LibreOffice); cai para LibreOffice automaticamente
    # se as credenciais não estiverem definidas, para nunca deixar o botão quebrado.
    engine_used = "aspose"
    try:
        pdf_path = _convert_with_aspose(src)
    except Exception as e:
        engine_used = "libreoffice"
        pdf_path = _convert_with_libreoffice(src)
        aviso = f"Gerado com LibreOffice (fallback) — Aspose indisponível: {e}"
    else:
        aviso = "Gerado com Aspose.Slides Cloud — fiel ao PowerPoint."

    return JSONResponse(
        {
            "status": "ok",
            "pdf_id": pptx_id,
            "pdf_url": f"{PUBLIC_BASE_URL}/files/pdf/{pptx_id}.pdf",
            "engine": engine_used,
            "aviso": aviso,
        }
    )


def _get_aspose_token() -> str | None:
    """Autentica na Aspose.Slides Cloud (OAuth2 client_credentials). None se não configurado."""
    client_id = os.environ.get("ASPOSE_CLIENT_ID")
    client_secret = os.environ.get("ASPOSE_CLIENT_SECRET")
    if not client_id or not client_secret:
        return None
    resp = httpx.post(
        "https://api.aspose.cloud/connect/token",
        data={"grant_type": "client_credentials", "client_id": client_id, "client_secret": client_secret},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def _convert_with_aspose(src_path: Path) -> Path:
    """
    Converte via Aspose.Slides Cloud — motor próprio (não usa LibreOffice), que
    lê corretamente as fontes incorporadas do arquivo e não tem o bug de
    sobreposição/corte que vimos no LibreOffice. Requer conta na Aspose Cloud
    (tem camada gratuita): https://dashboard.aspose.cloud -> crie um app ->
    defina ASPOSE_CLIENT_ID e ASPOSE_CLIENT_SECRET no Render.
    """
    token = _get_aspose_token()
    if token is None:
        raise RuntimeError("Aspose não configurado (faltam ASPOSE_CLIENT_ID/ASPOSE_CLIENT_SECRET)")

    with open(src_path, "rb") as f:
        resp = httpx.post(
            "https://api.aspose.cloud/v3.0/slides/convert/pdf",
            headers={"Authorization": f"Bearer {token}"},
            files={"file": (src_path.name, f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
            timeout=90,
        )
    resp.raise_for_status()

    pdf_path = PDF_DIR / f"{src_path.stem}.pdf"
    pdf_path.write_bytes(resp.content)
    return pdf_path


# Imagens sem recompressão JPEG e sem reduzir resolução (fotos e laterais ficam nítidas no PDF)
PDF_FILTER = ('pdf:impress_pdf_Export:{"ReduceImageResolution":{"type":"boolean","value":"false"},'
              '"UseLosslessCompression":{"type":"boolean","value":"true"}}')


def _convert_with_libreoffice(src_path: Path) -> Path:
    # O LibreOffice distribui a entrelinha de forma diferente do Canva/PowerPoint: o texto sai alguns pontos
    # deslocado. Convertemos uma CÓPIA ajustada (o .pptx original que o cliente baixa não é alterado).
    work_dir = Path(tempfile.mkdtemp(prefix="lo_"))
    tuned = work_dir / src_path.name
    try:
        mail_merge.tune_for_libreoffice(src_path, tuned)
    except Exception:
        shutil.copy(src_path, tuned)       # nunca falha o PDF por causa do ajuste fino
    try:
        result = subprocess.run(
            [os.environ.get("LIBREOFFICE_PATH", "soffice"), f"-env:UserInstallation=file:///tmp/lo_{uuid.uuid4().hex}",
             "--headless", "--convert-to", PDF_FILTER, "--outdir", str(PDF_DIR), str(tuned)],
            capture_output=True,
            text=True,
            timeout=180,
        )
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
    if result.returncode != 0:
        raise HTTPException(500, f"Falha ao converter para PDF: {result.stderr}")
    pdf_path = PDF_DIR / f"{src_path.stem}.pdf"
    if not pdf_path.exists():
        raise HTTPException(500, "Conversão terminou sem erro mas o PDF não foi encontrado.")
    return pdf_path


def _fonts_installed():
    """False => o PDF sairá com fonte trocada (letras espaçadas). Veja fonts/LEIAME.txt e o Dockerfile."""
    try:
        out = subprocess.run(["fc-list"], capture_output=True, text=True, timeout=10).stdout
        return all(name in out for name in ("Eastman Alternate Trial Bold", "Neulis Neue Bold"))
    except Exception:
        return None


@app.get("/health")
def health():
    template_found = TEMPLATE_PATH is not None and TEMPLATE_PATH.exists()
    field_map_found = FIELD_MAP_PATH is not None and FIELD_MAP_PATH.exists()
    
    if not template_found or not field_map_found:
        return JSONResponse(
            {
                "status": "error",
                "template_encontrado": template_found,
                "field_map_encontrado": field_map_found,
                "template_path_procurado": [str(p) for p in TEMPLATE_CANDIDATES],
                "field_map_path_procurado": [str(p) for p in FIELD_MAP_CANDIDATES],
                **({"template_error": TEMPLATE_ERROR} if TEMPLATE_PATH is None else {}),
                **({"field_map_error": FIELD_MAP_ERROR} if FIELD_MAP_PATH is None else {}),
            },
            status_code=503,
        )
    
    return JSONResponse(
        {
            "status": "ok",
            "storage_backend": os.environ.get("STORAGE_BACKEND", "local"),
            "fontes_do_template_instaladas": _fonts_installed(),
            "template_encontrado": template_found,
            "field_map_encontrado": field_map_found,
            "template_path": str(TEMPLATE_PATH),
            "field_map_path": str(FIELD_MAP_PATH),
        }
    )
