import json

from fastapi.testclient import TestClient

from app.main import app


def test_health_endpoint():
    client = TestClient(app)
    response = client.get("/health")
    assert response.status_code == 200
    payload = response.json()
    assert payload["status"] == "ok"
    assert payload["storage_backend"] in {"local", "s3"}


def test_generate_proposal_accepts_minimal_payload():
    client = TestClient(app)
    payload = {"NOME_ESCOLA": "Escola Teste"}
    response = client.post(
        "/api/propostas/gerar",
        data={
            "registro_id": "TEST-1",
            "dados": json.dumps(payload),
        },
    )
    assert response.status_code == 200
    body = response.json()
    assert body["status"] in {"ok", "ok_com_avisos"}
    assert body["pptx_url"].startswith("/files/") or body["pptx_url"].startswith("http")


def test_generate_pdf_missing_returns_404():
    client = TestClient(app)
    response = client.post("/api/propostas/pdf", data={"pptx_id": "missing-id"})
    assert response.status_code == 404
    assert "não encontrado" in response.json()["detail"]
