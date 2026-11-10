"""Dívida maior que a operação escolar: modos via_cash_in (padrão), imovel_absorve e proporcional + travas."""
import pytest
from decimal import Decimal
from app import mail_merge as mm


def _run(esc, imo, div, **extra):
    d = {"VALOR_OPERACAO_ESCOLAR": f"{esc}".replace(".", ","), "VALOR_PATRIMONIO_IMOBILIARIO": f"{imo}".replace(".", ","),
         "VALOR_ASSUNCAO_DIVIDA": f"{div}".replace(".", ","), "ASSUNCAO_DIVIDA_ATIVA": "1",
         "PERCENTUAL_CASH_OUT": "60%", "PERCENTUAL_CASH_IN": "40%", "NOME_ESCOLA": "COLEGIO X"}
    d.update(extra)
    rep = mm.new_report()
    mm.apply_cash_rule(d, rep)
    return d, rep


def _v(s):
    return Decimal(s.replace("R$", "").replace(".", "").replace(",", ".").strip())


def test_padrao_via_cash_in_mantem_cash_out_dos_vendedores():
    d, _ = _run(300000, 700000, 400000)             # D <= 40% do subtotal
    assert d["MODO_DIVIDA_APLICADO"] == "via_cash_in"
    assert _v(d["ESCOLA_TOTAL_DESTINADO_VENDEDOR"]) == Decimal("180000.00")   # 60% da escola, intacto
    assert _v(d["IMOVEL_TOTAL_DESTINADO_VENDEDOR"]) == Decimal("420000.00")
    assert _v(d["VALOR_CASH_IN"]) == Decimal("0.00")          # 40% do subtotal (400k) - dívida (400k)
    d, _ = _run(300000, 700000, 350000)
    assert _v(d["VALOR_CASH_IN"]) == Decimal("50000.00") and _v(d["VALOR_CASH_OUT"]) == Decimal("600000.00")


def test_proporcional_quando_via_cash_in_inviavel():
    d, _ = _run(900000, 600000, 1100000)
    assert d["MODO_DIVIDA_APLICADO"] == "proporcional"
    assert _v(d["ESCOLA_TOTAL_DESTINADO_VENDEDOR"]) + _v(d["IMOVEL_TOTAL_DESTINADO_VENDEDOR"]) == Decimal("240000.00")
    assert _v(d["ESCOLA_TOTAL_DESTINADO_VENDEDOR"]) == Decimal("144000.00")


def test_escola_pequena_dívida_no_imovel_quando_cash_in_inviavel():
    d, _ = _run(50000, 600000, 300000)
    assert d["MODO_DIVIDA_APLICADO"] == "imovel_absorve"
    assert _v(d["ESCOLA_TOTAL_DESTINADO_VENDEDOR"]) == Decimal("30000.00")


def test_modo_forcado_e_divida_menor_que_escola_segue_original():
    d, _ = _run(300000, 700000, 100000)
    assert d["MODO_DIVIDA_APLICADO"] == "original"
    d, _ = _run(300000, 700000, 100000, MODO_DIVIDA="proporcional")
    assert d["MODO_DIVIDA_APLICADO"] == "proporcional"


def test_bloqueio_divida_maior_que_subtotal():
    with pytest.raises(ValueError):
        _run(300000, 700000, 1000000)


def test_trava_30_pct_gera_aviso():
    _, rep = _run(900000, 600000, 1400000, MODO_DIVIDA="proporcional")
    assert any("mínimo de 30%" in x for x in rep["inconsistencias"])
