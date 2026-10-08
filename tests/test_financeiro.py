"""Slides 6 e 7: nenhuma quantidade de parcelas nem valor pode ficar fixo; singular/plural pela quantidade."""
import json
from pathlib import Path
from pptx import Presentation
from app import mail_merge as mm

ROOT = Path(__file__).resolve().parent.parent
BASE = dict(
    CURTO_PROLABORE_MENSAL="R$ 7.000,00", MEDIO_PROLABORE_MENSAL="R$ 7.000,00", LONGO_PROLABORE_MENSAL="R$ 7.000,00",
    CURTO_BONUS_AQUISICAO="R$ 3.000,00", MEDIO_ALUGUEL_TRANSITORIO_MENSAL="R$ 3.000,00",
    PCT_PARTICIPACAO_MANTENEDOR="30%", PCT_PARTICIPACAO_ADQUIRIDA="70%", CURTO_TOTAL_MENSAL="R$ 10.000,00",
)
SAGRADA = dict(BASE, ESCOLA_QTD_PARCELAS_ENTRADA="1", ESCOLA_QTD_PARCELAS_SALDO="48",
               IMOVEL_QTD_PARCELAS_ENTRADA="1", IMOVEL_QTD_PARCELAS_SALDO="48",
               MEDIO_PARCELA_ENTRADA_NEGOCIO="R$ 40.941,57", LONGO_PARCELA_SALDO_NEGOCIO="R$ 7.676,55",
               LONGO_PARCELA_ENTRADA_IMOVEL="R$ 145.592,45", IMOVEL_VALOR_PARCELA_SALDO="R$ 27.298,58",
               MEDIO_TOTAL_MENSAL="R$ 50.941,57", LONGO_TOTAL_MENSAL="R$ 160.269,00")
CRISTO = dict(SAGRADA, ESCOLA_QTD_PARCELAS_ENTRADA="3", ESCOLA_QTD_PARCELAS_SALDO="60",
              IMOVEL_QTD_PARCELAS_ENTRADA="3", IMOVEL_QTD_PARCELAS_SALDO="60",
              MEDIO_PARCELA_ENTRADA_NEGOCIO="R$ 9.712,93", LONGO_PARCELA_SALDO_NEGOCIO="R$ 2.708,98",
              LONGO_PARCELA_ENTRADA_IMOVEL="R$ 56.959,74", IMOVEL_VALOR_PARCELA_SALDO="R$ 15.886,31",
              MEDIO_TOTAL_MENSAL="R$ 19.712,93", LONGO_TOTAL_MENSAL="R$ 69.668,72")


def _texts(data, tmp_path, slide):
    out = tmp_path / "o.pptx"
    mm.generate(ROOT / "modelo_rede_lumo (1).pptx", mm.load_json(ROOT / "field_map.json"), data, out)
    prs = Presentation(str(out))
    res = {}
    for sh, *_ in mm._iter_shapes_abs(prs.slides[slide - 1].shapes):
        if sh.has_text_frame:
            res[sh.shape_id] = sh.text_frame.text.replace("\x0b", " ")
    return res


def test_sagrada_familia_slide7(tmp_path):
    t = _texts(SAGRADA, tmp_path, 7)
    assert t[98] == "1 parcela de R$ 40.941,57"
    assert t[91] == "Após a parcela inicial, o saldo do negócio passa para 48 parcelas de R$ 7.676,55."
    assert t[100] == "48 parcelas de R$ 7.676,55" and t[99] == "1 parcela de R$ 145.592,45"
    assert t[92] == ("Após a parcela inicial, os saldos passam para 48 parcelas de R$ 7.676,55 no negócio "
                     "e 48 parcelas de R$ 27.298,58 no imóvel.")
    assert t[110] == t[111] == "Total mensal na parcela inicial"
    assert (t[101], t[102], t[103]) == ("R$ 10.000,00", "R$ 50.941,57", "R$ 160.269,00")
    assert t[117].startswith("30% de participação sobre 70%")
    joined = " ".join(t.values())
    assert "3 parcelas" not in joined and "60 parcelas" not in joined and "2.708,98" not in joined


def test_cristo_rei_slide7(tmp_path):
    t = _texts(CRISTO, tmp_path, 7)
    assert t[98] == "3 parcelas de R$ 9.712,93" and t[100] == "60 parcelas de R$ 2.708,98"
    assert t[110] == "Total mensal nas 3 parcelas iniciais"
    assert t[103] == "R$ 69.668,72"


def test_slide6_singular(tmp_path):
    t = _texts(SAGRADA, tmp_path, 6)
    assert t[34].startswith("1 primeiro pagamento") and "Pago em 1 parcela" in t[34]
    assert t[72].startswith("1 primeiro pagamento") and "1 primeiros" not in t[72]
    assert t[31].startswith("em 1 parcela de") and t[33].startswith("em 48 parcelas de")
    assert t[35].startswith("Saldo em 48 meses")
