# Backend — Geração de Propostas Rede Lumo

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/rodrigomonte07/Backend-lumo)

Implementa os dois endpoints descritos no prompt do Lovable. Testado ponta a ponta (gerar → pdf) com sucesso.

## ⚡ Quick Start

### Rodar localmente

```bash
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

**Pré-requisitos:**
- Python 3.8+
- LibreOffice (para PDF): `apt install libreoffice` (Ubuntu/Debian) ou `brew install libreoffice` (macOS)

### Rodar com Docker

```bash
docker build -t rede-lumo-backend .
docker run -p 8000:8000 -v $(pwd)/storage:/app/storage rede-lumo-backend
```

### Deploy em Render

Clique no botão acima ou acesse [render.com](https://render.com) e:

1. Conecte seu repositório GitHub
2. Crie um novo serviço web
3. Use o comando: `uvicorn app.main:app --host 0.0.0.0 --port $PORT`
4. Adicione um disco persistente para `/app/storage` (1GB mínimo)
5. Garanta que a variável `LIBREOFFICE_PATH` esteja definida como `/usr/bin/soffice`

## 📋 Endpoints

### POST `/api/propostas/gerar`

Gera uma proposta em .pptx a partir do template + dados do CRM.

**Request:**
```bash
curl -X POST http://localhost:8000/api/propostas/gerar \
  -F "registro_id=ESCOLA_123" \
  -F "dados=@dados.json" \
  -F "foto_fachada=@fachada.jpg"
```

**dados.json:**
```json
{
  "NOME_ESCOLA": "Escola Vida de Criança",
  "AREA_CONSTRUIDA": "820 m²",
  "RECEITA_ANUAL": "R$ 406.159,53",
  "ALUNOS_ATUAIS": "105",
  "SALAS_ATIVAS": "11",
  "PERCENTUAL_OCUPACAO": "41,34%"
}
```

**Response:**
```json
{
  "status": "ok",
  "pptx_id": "c6b586ff...",
  "pptx_url": "/files/pptx/c6b586ff....pptx",
  "registro_id": "ESCOLA_123",
  "campos_aplicados": 68,
  "avisos": {
    "tokens_faltando": [],
    "regras_nao_encontradas": [],
    "erros": []
  }
}
```

**Form Fields:**
- `registro_id` (text, required) — ID do registro no CRM para rastreio
- `dados` (text/JSON, required) — Dicionário token → valor (ver `field_map.json` para lista completa)
- `foto_fachada` (file, optional) — Foto da escola

**Respostas:**
- `200 OK` — Proposta gerada com sucesso
- `400 Bad Request` — JSON inválido em `dados`
- `500 Server Error` — Template ou field_map não encontrados

---

### POST `/api/propostas/pdf`

Converte um .pptx já gerado em .pdf.

**Request:**
```bash
curl -X POST http://localhost:8000/api/propostas/pdf \
  -F "pptx_id=c6b586ff..."
```

**Response:**
```json
{
  "status": "ok",
  "pdf_id": "c6b586ff...",
  "pdf_url": "/files/pdf/c6b586ff....pdf",
  "aviso": "Convertido com LibreOffice — valide visualmente antes de produção."
}
```

**Form Fields:**
- `pptx_id` (text, required) — ID retornado por `/api/propostas/gerar`

**Respostas:**
- `200 OK` — PDF convertido com sucesso
- `404 Not Found` — pptx_id não encontrado
- `500 Server Error` — LibreOffice não instalado ou timeout

---

### GET `/health`

Health check para orchestration/deployment.

**Response:**
```json
{
  "status": "ok",
  "environment": "production",
  "template_encontrado": true,
  "field_map_encontrado": true
}
```

**Retorna 503 Service Unavailable** se template ou field_map não forem encontrados.

---

## 📁 Arquivos

```
.
├── app/
│   ├── __init__.py           # Package init
│   ├── config.py             # Configurações (env vars, paths)
│   ├── main.py               # Endpoints FastAPI
│   └── mail_merge.py         # Motor de substituição PPTX
├── field_map.json            # Mapa de 68 campos
├── modelo_rede_lumo (1).pptx # Template original (Canva)
├── storage/                  # Gerado em runtime
│   ├── pptx/                 # PPTXs gerados
│   ├── pdf/                  # PDFs convertidos
│   └── uploads/              # Fotos enviadas
├── main.py                   # Entrypoint raiz
├── requirements.txt          # Deps Python
├── Dockerfile                # Build Docker
├── render.yaml               # Deploy Render
├── README.md                 # Este arquivo
└── .gitignore                # Arquivos locais ignorados
```

> Importante: no ambiente atual, o template e o `field_map.json` ficam na raiz do projeto, e a aplicação já resolve esses caminhos automaticamente.

## 🔧 Configuração via Variáveis de Ambiente

```bash
# Ambiente
ENVIRONMENT=production              # development | production
LOG_LEVEL=INFO                      # DEBUG | INFO | WARNING | ERROR

# LibreOffice
LIBREOFFICE_PATH=/usr/bin/soffice  # Caminho do binário (default: soffice)
PDF_TIMEOUT=90                      # Timeout em segundos

# CORS
CORS_ORIGINS=*                      # Origens permitidas (comma-separated)
```

## ⚠️ Avisos de Produção

### 1. LibreOffice / Conversão PDF

Este endpoint usa LibreOffice headless, que tem um bug de renderização **confirmado** neste template específico:
- Slide "Valoração da Transação": sobreposição de texto e corte de dígito
- O arquivo `.pptx` do `/gerar` está sempre correto
- O risco é só na conversão automática para PDF

**Antes de produção:**
Troque `_convert_with_libreoffice()` em `app/main.py` por um serviço mais fiel:
- **Aspose.Slides Cloud** — Excelente renderização
- **Adobe PDF Services API** — Profissional
- **CloudConvert** — Com engine PowerPoint/Office

A interface da função é simples (caminho `.pptx` entra, caminho `.pdf` sai), então é uma troca de poucas linhas.

### 2. Armazenamento

**Local (atual):**
- Arquivos ficam em `storage/` no disco local
- Não persiste em ambientes serverless (Render free, Vercel, etc)

**Produção recomendada:**
- **AWS S3**
- **Google Cloud Storage**
- **Azure Blob Storage**

Para trocar: edite `app/main.py`, linhas onde faz `prs.save(str(out_path))` e `mail_merge.replace_photo()`, e use SDK do seu cloud provider.

### 3. Dados Incompletos

Se `avisos.tokens_faltando` não estiver vazio:
- O token existe em `field_map.json` mas não veio em `dados`
- O campo correspondente no slide fica com o texto original do template
- Trate como aviso **não-bloqueante** (ou bloqueie no front, a seu critério)

### 4. Performance

- Conversão PDF leva ~10-15s (timeout default: 90s)
- Em produção, monitore logs em `/health` e em endpoints

## ✅ O que já foi testado

- ✅ Geração completa com os 68 campos + foto → sucesso
- ✅ Conversão para PDF → sucesso (arquivo gerado, ressalva do LibreOffice acima)
- ✅ Dados incompletos → não trava, retorna aviso com tokens faltando
- ✅ JSON malformado → erro 400 claro
- ✅ pptx_id inexistente → erro 404 claro
- ✅ LibreOffice ausente → erro 500 informativo

## 🚀 Próximos Passos

1. **Local:**
   ```bash
   pip install -r requirements.txt
   uvicorn app.main:app --reload
   ```

2. **Teste com curl:**
   ```bash
   curl -X POST http://localhost:8000/api/propostas/gerar \
     -F "registro_id=TEST" \
     -F "dados={\"NOME_ESCOLA\":\"Teste\"}"
   ```

3. **Deploy em Render:**
   - Clique no botão Deploy acima
   - Ou conecte manualmente: https://dashboard.render.com

4. **Substitua a conversão PDF:**
   - Se usar em produção real, troque LibreOffice conforme ⚠️ acima

## 📞 Troubleshooting

### "field_map.json não encontrado"
- Verifique se existe na raiz do projeto ou em `template/`
- Leia o erro do `/health` endpoint

### "Template .pptx não encontrado"
- Verifique se existe na raiz do projeto ou em `template/`
- Pode estar em `modelo_rede_lumo.pptx` ou `modelo_rede_lumo (1).pptx`

### "LibreOffice não encontrado (PDF)"
- Em container: já vem instalado no Dockerfile
- Local: `apt install libreoffice` (Linux) ou `brew install libreoffice` (macOS)

### Timeout ao converter PDF
- LibreOffice lento: ajuste `PDF_TIMEOUT` env var
- Máquina sobrecarregada: considere usar serviço em nuvem

## 📄 Licença

MIT

---

**Status:** ✅ Pronto para produção (com ressalvas de PDF acima)


## Fidelidade ao modelo do Canva (como funciona)

O gerador parte do .pptx original e usa o PDF exportado do Canva (`assets/modelo_canva_referencia.pdf`) como gabarito:

| Arquivo em `assets/` | O que guarda | Gerado por |
|---|---|---|
| `edges/` | laterais (e cantos da capa/fecho) como UMA imagem por lado | `tools/build_edges.py` |
| `pills.json` | fundos arredondados atrás de textos (o Canva não exporta no .pptx) | `tools/build_pills.py` |
| `line_breaks.json` | quebras de linha exatas do Canva | `tools/build_breaks.py` |
| `line_pitch.json` | espaçamento real entre linhas | `tools/build_pitch.py` |
| `lo_offsets.json` | ajuste vertical só para a conversão em PDF (LibreOffice) | `tools/build_lo_offsets.py` |

**Se trocar o template/PDF do Canva**, rode nesta ordem (precisa de `soffice`, `poppler-utils` e `pip install -r requirements-tools.txt`):
`build_edges.py`, `build_pills.py`, `build_breaks.py`, `build_pitch.py`, `build_lo_offsets.py`.

Slide 2: a grade de 87 bonequinhos e os colchetes seguem a % de ocupação (`PERCENTUAL_OCUPACAO`).
As duas "pizzas" são ilustrações fixas do modelo (um quarto destacado), não gráficos de dados.

## Slides 6 e 7: parcelas e valores (nada fixo)

Quantidades e valores vêm de `dados`; singular/plural pela quantidade ("1 parcela" / "N parcelas").
- Quantidades: `ESCOLA_QTD_PARCELAS_ENTRADA`, `ESCOLA_QTD_PARCELAS_SALDO`, `IMOVEL_QTD_PARCELAS_ENTRADA`, `IMOVEL_QTD_PARCELAS_SALDO`
- Valores: `MEDIO_PARCELA_ENTRADA_NEGOCIO`, `LONGO_PARCELA_SALDO_NEGOCIO`, `LONGO_PARCELA_ENTRADA_IMOVEL`, `IMOVEL_VALOR_PARCELA_SALDO`
- Totais (usados como vêm, sem recálculo): `CURTO_TOTAL_MENSAL`, `MEDIO_TOTAL_MENSAL`, `LONGO_TOTAL_MENSAL`
- Participação: `PCT_PARTICIPACAO_MANTENEDOR`, `PCT_PARTICIPACAO_ADQUIRIDA`
- Título do total: "Total mensal na parcela inicial" (1) ou "Total mensal nas N parcelas iniciais" (N>1): médio usa a entrada da escola, longo a entrada do imóvel.
- Se faltar algum campo, o texto correspondente não é gerado e aparece em `inconsistencias` no retorno (o texto do modelo não deve ser usado nesse caso).
