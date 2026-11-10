"""Slide 7: totais batem com as linhas visíveis; faixa de adiantamento (13º) só para quem tem."""
import json, sys
from pathlib import Path
from pptx import Presentation
from app import mail_merge as mm
sys.path.insert(0, str(Path(__file__).parent))
import test_escola_zerada as base

DADOS = dict(base.DADOS, NOME_ESCOLA="ANDERSON ALVES COLEGIO", VALOR_OPERACAO_ESCOLAR="R$ 59.202,36", VALOR_TOTAL_TRANSACAO="R$ 928.611,00",
             VALOR_ASSUNCAO_DIVIDA="R$ 63.000,00", IMOVEL_VALOR_PARCELA_SALDO="R$ 9.072,10",
             MEDIO_PARCELA_ENTRADA_NEGOCIO="R$ 3.552,14", LONGO_PARCELA_SALDO_NEGOCIO="R$ 666,03",
             LONGO_PARCELA_ENTRADA_IMOVEL="R$ 48.384,52", CURTO_TOTAL_MENSAL="R$ 8.000,00",
             MEDIO_TOTAL_MENSAL="R$ 11.311,15", LONGO_TOTAL_MENSAL="R$ 54.246,35")


def _t7(dados, tmp):
    out = tmp / "o.pptx"
    mm.generate(base.TEMPLATE, mm.load_json(base.ROOT / "field_map.json"), dados, out)
    return base._texts(Presentation(str(out)), 7)


def test_totais_batem_com_linhas(tmp_path):
    t = _t7(dict(DADOS), tmp_path)
    assert "R$ 11.552,14" in t and "R$ 52.250,55" in t and "R$ 8.000,00" in t
    assert "11.311,15" not in t and "54.246,35" not in t


def test_adiantamento_so_anderson(tmp_path):
    t = _t7(dict(DADOS), tmp_path)
    assert "13º salário" in t and "R$ 14.000,00" in t and "R$ 16.000,00" in t and "10/12" in t
    t2 = _t7(dict(DADOS, NOME_ESCOLA="OUTRA ESCOLA"), tmp_path)
    assert "13º" not in t2
    t3 = _t7(dict(DADOS, NOME_ESCOLA="OUTRA ESCOLA", ADIANTAMENTO_1_VALOR="R$ 1.000,00", ADIANTAMENTO_1_PRAZO="hoje"), tmp_path)
    assert "R$ 1.000,00" in t3


def test_adiantamento_abate_imovel_e_nao_escola(tmp_path):
    out = tmp_path / "o.pptx"
    mm.generate(base.TEMPLATE, mm.load_json(base.ROOT / "field_map.json"), dict(DADOS), out)
    prs = Presentation(str(out))
    t6, t5 = base._texts(prs, 6), base._texts(prs, 5)
    assert "R$ 35.521,42" in t6            # escola intacta
    assert "R$ 465.845,18" in t6           # imóvel: (869.408,64 - 63.000 - 30.000) x 60%
    assert "R$ 501.366,60" in t6 and "Adiantamento (13º salário)" in t5
