"""Variante sem dívidas / sem concorrência (slides 3 e 8) — só para a escola marcada."""
from pathlib import Path
from pptx import Presentation
from app import mail_merge as mm

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "modelo_rede_lumo (1).pptx"


def _texts(prs, slide):
    return " | ".join(s.text_frame.text for s in prs.slides[slide - 1].shapes if s.has_text_frame)


def _gerar(tmp_path, dados):
    out = tmp_path / "o.pptx"
    mm.generate(TEMPLATE, mm.load_json(ROOT / "field_map.json"), dados, out)
    return Presentation(str(out))


def test_cesp_slides_3_e_8(tmp_path):
    prs = _gerar(tmp_path, {"NOME_ESCOLA": "Cesp - Centro de Educacao Souza Pimentel"})
    t3, t8 = _texts(prs, 3), _texts(prs, 8)
    assert "299" in t3 and "escolas concorrentes" in t3 and "+24" in t3
    assert "3.800" not in t3 and "304" not in t3 and "14" not in t3.split("|")[0]
    assert "Sem passivos identificados" in t8 and "130.940,86" not in t8


def test_outra_escola_nao_muda(tmp_path):
    prs = _gerar(tmp_path, {"NOME_ESCOLA": "Outra Escola", "RESERVA_DIVIDAS_VALOR": "R$ 10.000,00"})
    assert "Sem passivos" not in _texts(prs, 8) and "10.000,00" in _texts(prs, 8)
    assert "escolas concorrentes" not in _texts(prs, 3)


def test_flag_crm_e_crescimento_customizado(tmp_path):
    prs = _gerar(tmp_path, {"NOME_ESCOLA": "X", "MODO_PROPOSTA": "sem_dividas_sem_concorrencia",
                            "ALUNOS_REGIAO": "299", "CRESCIMENTO_5ANOS_PCT": "10%"})
    assert "+30" in _texts(prs, 3) and "Sem passivos identificados" in _texts(prs, 8)
