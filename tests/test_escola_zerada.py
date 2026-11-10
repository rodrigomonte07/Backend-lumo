"""Dívida maior que a operação escolar: escola zera, excedente abate o imóvel, sem valores negativos."""
import json
from pathlib import Path
from pptx import Presentation
from app import mail_merge as mm

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "modelo_rede_lumo (1).pptx"
DADOS = {
    "NOME_ESCOLA": "COLEGIO TESTE DIVIDA", "PERCENTUAL_PARTICIPACAO_SOCIETARIA": "70%",
    "VALOR_OPERACAO_ESCOLAR": "R$ 20.000,00", "VALOR_PATRIMONIO_IMOBILIARIO": "R$ 869.408,64",
    "VALOR_TOTAL_TRANSACAO": "R$ 889.408,64", "PERCENTUAL_CASH_OUT": "60%", "PERCENTUAL_CASH_IN": "40%",
    "ASSUNCAO_DIVIDA_ATIVA": "sim", "VALOR_ASSUNCAO_DIVIDA": "R$ 100.000,00",
    "PERCENTUAL_ENTRADA": "10%", "ESCOLA_QTD_PARCELAS_ENTRADA": "1", "ESCOLA_QTD_PARCELAS_SALDO": "48",
    "IMOVEL_QTD_PARCELAS_ENTRADA": "1", "IMOVEL_QTD_PARCELAS_SALDO": "48",
    "CURTO_PROLABORE_MENSAL": "R$ 5.000,00", "MEDIO_PROLABORE_MENSAL": "R$ 5.000,00", "LONGO_PROLABORE_MENSAL": "R$ 5.000,00",
    "CURTO_BONUS_AQUISICAO": "R$ 3.000,00", "MEDIO_ALUGUEL_TRANSITORIO_MENSAL": "R$ 3.000,00",
    "IMOVEL_VALOR_PARCELA_SALDO": "R$ 9.780,85", "PCT_PARTICIPACAO_MANTENEDOR": "30%", "PCT_PARTICIPACAO_ADQUIRIDA": "70%",
}


def _texts(prs, n):
    return " | ".join(s.text_frame.text for s in prs.slides[n - 1].shapes if s.has_text_frame)


def test_divida_maior_que_escola_abate_o_imovel(tmp_path):
    out = tmp_path / "o.pptx"
    mm.generate(TEMPLATE, mm.load_json(ROOT / "field_map.json"), dict(DADOS), out)
    prs = Presentation(str(out))
    for n in (5, 6, 7):
        assert "R$ -" not in _texts(prs, n)
    t6 = _texts(prs, 6)
    assert "R$ 12.000,00" in t6          # escola: 20.000 x 60%, sem abatimento
    assert "R$ 461.645,18" in t6         # imóvel: (869.408,64 - 100.000) x 60%
    assert "R$ 473.645,18" in t6         # total igual ao do CRM
    assert "NEGOCIAÇÃO DA ESCOLA" in t6
