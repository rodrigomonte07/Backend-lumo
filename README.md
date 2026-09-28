# Backend — Geração de Propostas Rede Lumo

Implementa os dois endpoints descritos no prompt do Lovable. Testado
ponta a ponta (gerar → pdf) com sucesso.

## Rodar localmente
```bash
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```
Precisa também do `soffice` (LibreOffice) instalado no sistema para o
endpoint de PDF — no Ubuntu/Debian: `apt install libreoffice`.

## Endpoints

### POST /api/propostas/gerar
`multipart/form-data`:
- `registro_id` (texto) — id do registro no seu CRM, só para rastreio
- `dados` (texto, JSON) — dicionário token → valor (ver `template/field_map.json`
  para a lista completa de tokens esperados)
- `foto_fachada` (arquivo, opcional) — foto da escola

Resposta:
```json
{
  "status": "ok",
  "pptx_id": "c6b586ff...",
  "pptx_url": "/files/pptx/c6b586ff....pptx",
  "campos_aplicados": 68,
  "avisos": { "tokens_faltando": [], "regras_nao_encontradas": [], "erros": [] }
}
```
Se `avisos.tokens_faltando` não estiver vazio, o token existe no field_map
mas não veio em `dados` — o campo correspondente no slide fica com o texto
antigo do template. Trate isso no front como um aviso não-bloqueante (ou
bloqueie a geração, a seu critério).

### POST /api/propostas/pdf
`multipart/form-data` ou `application/x-www-form-urlencoded`:
- `pptx_id` — o id retornado por `/gerar`

Resposta:
```json
{ "status": "ok", "pdf_id": "c6b586ff...", "pdf_url": "/files/pdf/c6b586ff....pdf" }
```

**⚠️ Antes de produção:** este endpoint usa LibreOffice, que tem um bug de
renderização confirmado neste template específico (ver conversa anterior —
slide "Valoração da Transação"). O arquivo abre perfeito no PowerPoint/Canva,
então o `.pptx` do endpoint `/gerar` está sempre correto — o risco é só na
conversão automática para PDF. Troque `_convert_with_libreoffice()` em
`app/main.py` por um serviço mais fiel antes de ir ao ar (Aspose.Slides
Cloud, Adobe PDF Services, CloudConvert). A assinatura da função
(caminho do .pptx entra, caminho do .pdf sai) já está isolada para isso
ser uma troca de poucas linhas.

## Arquivos
- `app/main.py` — os dois endpoints + servidor de arquivos estáticos.
- `app/mail_merge.py` — o motor de substituição (mesma lógica já validada).
- `template/modelo_rede_lumo.pptx` — o arquivo original do Canva. NÃO EDITE.
- `template/field_map.json` — as 68 regras de substituição.
- `storage/` — pptx/pdf gerados e fotos enviadas (troque por S3/GCS em produção;
  hoje fica em disco local, que não persiste em ambientes serverless).

## O que já foi testado
- Geração completa com os 68 campos + foto → sucesso.
- Conversão para PDF → sucesso (arquivo gerado, ressalva do LibreOffice acima).
- Dados incompletos → não trava, retorna aviso com os tokens faltando.
- JSON malformado em `dados` → erro 400 claro.
- `pptx_id` inexistente em `/pdf` → erro 404 claro.
