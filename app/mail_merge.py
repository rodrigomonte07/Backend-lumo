#!/usr/bin/env python3
"""
Mail-merge engine para o template "Rede Lumo" (gerado no Canva).

Em vez de pedir para uma IA "redesenhar" o slide, este script abre o
ARQUIVO ORIGINAL do Canva e troca apenas o conteúdo (texto e foto) nos
pontos exatos mapeados em field_map.json — a formatação (fonte, cor,
posição, forma) nunca é tocada, porque nunca é recriada.

Uso:
    python3 mail_merge.py --template modelo.pptx --data data.json \
        --map field_map.json --foto fachada.jpg --out proposta_final.pptx

data.json: dicionário token -> valor (string já formatada, ex:
    {"NOME_ESCOLA": "Escola Vida de Criança", "AREA_CONSTRUIDA": "820 m²", ...}
Todos os tokens usados em field_map.json devem existir em data.json.
"""

import argparse
import json
import math
import re
import shutil
import sys
import zipfile
from io import BytesIO
from functools import lru_cache
from pathlib import Path

from PIL import Image, ImageFilter, ImageFont
from pptx import Presentation
from pptx.util import Emu, Pt

EMU_PER_PT = 12700
BASE_DIR = Path(__file__).resolve().parent.parent
FONTS_DIR = BASE_DIR / "fonts"
EDGES_DIR = BASE_DIR / "assets" / "edges"
PILLS_JSON = BASE_DIR / "assets" / "pills.json"
PITCH_JSON = BASE_DIR / "assets" / "line_pitch.json"
BREAKS_JSON = BASE_DIR / "assets" / "line_breaks.json"
LO_OFFSETS_JSON = BASE_DIR / "assets" / "lo_offsets.json"

# Faixas decorativas (laterais) são trocadas por UMA imagem por lado — ver tools/build_edges.py
EDGE_LEFT_LIMIT_EMU = 800000
EDGE_RIGHT_LIMIT_EMU = 16900000


# ----------------------------------------------------------------------------
# Medição de texto com as FONTES REAIS do template (pasta fonts/)
# ----------------------------------------------------------------------------
# Antes o ajuste de tamanho usava uma largura média de caractere "chutada", que
# errava para as fontes decorativas do Canva (Eastman Alternate / Neulis Neue).
# Agora medimos cada palavra com o arquivo da fonte de verdade (Pillow), então a
# decisão "cabe / não cabe" é a mesma que o PowerPoint tomaria.

MIN_SHRINK_RATIO = 0.60       # nunca encolhe a fonte abaixo de 60% do original
LINE_HEIGHT_FACTOR = 1.22     # entrelinha aproximada quando não há lnSpc
FIT_TOLERANCE = 0.93          # 7% de folga: LibreOffice/PowerPoint/Canva diferem alguns % na medição
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
DEFAULT_TYPEFACE = "Neulis Neue"


def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


@lru_cache(maxsize=64)
def _font_file(typeface):
    key = re.sub(r"[\s\-]", "", typeface or "")
    for ext in (".otf", ".ttf"):
        p = FONTS_DIR / f"{key}{ext}"
        if p.exists():
            return str(p)
    return None


@lru_cache(maxsize=64)
def _pil_font(path):
    return ImageFont.truetype(path, 1000)   # medimos em 1000 e escalamos


def _run_typeface(run):
    r = run._r
    rpr = r.find(f"{{{A_NS}}}rPr")
    if rpr is not None:
        lat = rpr.find(f"{{{A_NS}}}latin")
        if lat is not None and lat.get("typeface"):
            return lat.get("typeface")
    return DEFAULT_TYPEFACE


def _run_spacing_pt(run):
    rpr = run._r.find(f"{{{A_NS}}}rPr")
    if rpr is not None and rpr.get("spc"):
        try:
            return int(rpr.get("spc")) / 100.0
        except ValueError:
            return 0.0
    return 0.0


def _text_width_pt(text, typeface, size_pt, spacing_pt=0.0):
    path = _font_file(typeface)
    if path is None:                      # fonte desconhecida: estimativa conservadora
        return len(text) * size_pt * 0.62
    return _pil_font(path).getlength(text) / 1000.0 * size_pt + len(text) * spacing_pt


def _segments(runs_info, scale=1.0):
    """runs_info: [(text, typeface, size_pt, spacing_pt)] -> lista de palavras com largura."""
    words, cur = [], None
    for text, face, size, spc in runs_info:
        size = size * scale
        for piece in re.split(r"( |\n)", text):
            if piece == "":
                continue
            if piece == "\n":
                words.append(("\n", 0.0))
                cur = None
                continue
            w = _text_width_pt(piece, face, size, spc)
            if piece == " ":
                words.append((" ", w))
                cur = None
            elif cur is not None:
                words[-1] = (words[-1][0] + piece, words[-1][1] + w)   # palavra que atravessa runs
            else:
                words.append((piece, w))
                cur = True
    return words


def _count_lines(words, box_width_pt):
    """Quebra gulosa por palavra (igual ao PowerPoint). Palavra maior que a caixa quebra por letra."""
    lines, cur = 1, 0.0
    max_word = 0.0
    for text, w in words:
        if text == "\n":
            lines += 1
            cur = 0.0
            continue
        if text == " ":
            cur += w
            continue
        max_word = max(max_word, w)
        if w > box_width_pt:
            if cur > 0:
                lines += 1
            lines += int(w // box_width_pt)
            cur = w % box_width_pt
            continue
        if cur + w > box_width_pt and cur > 0:
            lines += 1
            cur = w
        else:
            cur += w
    return lines, max_word


def _inner_width_pt(shape):
    tf = shape.text_frame
    left = tf.margin_left if tf.margin_left is not None else 91440
    right = tf.margin_right if tf.margin_right is not None else 91440
    return max(1.0, (shape.width - left - right) / EMU_PER_PT)


def _para_runs_info(para, texts=None):
    """Lista (texto, fonte, tamanho, espaçamento) na ordem do parágrafo; <a:br/> vira o texto "\n"."""
    runs = list(para.runs)
    info, i, last_size = [], 0, 18.0
    for child in para._p:
        tag = child.tag.split("}")[1]
        if tag == "r" and i < len(runs):
            run = runs[i]
            size = run.font.size.pt if run.font.size is not None else 18.0
            t = texts[i] if texts is not None else run.text
            info.append((t, _run_typeface(run), size, _run_spacing_pt(run)))
            last_size, i = size, i + 1
        elif tag == "br":
            info.append(("\n", DEFAULT_TYPEFACE, last_size, 0.0))
    return info


def _is_nowrap(shape):
    bp = shape.text_frame._txBody.find(f"{{{A_NS}}}bodyPr")
    return bp is not None and bp.get("wrap") == "none"


def _para_alignment(para):
    ppr = para._p.find(f"{{{A_NS}}}pPr")
    return ppr.get("algn") if ppr is not None else None


def _fit_paragraph_inner(shape, para, original_texts, slide, report):
    """
    Garante que o parágrafo (já com o texto novo) ocupe no máximo o mesmo nº de
    linhas que o texto de exemplo do template ocupava.
      1) encolhe a fonte (até 60%)
      2) se ainda não couber e era 1 linha -> ALARGA a caixa (mantendo o centro se
         for texto centralizado) em vez de quebrar linha
      3) senão cresce a altura e empurra o que está logo abaixo
    """
    runs = list(para.runs)
    if not runs or shape.width in (None, 0):
        return
    box_w = _inner_width_pt(shape) * FIT_TOLERANCE

    # Caixas com wrap="none" crescem sozinhas na horizontal: só garantimos que não saiam da lâmina.
    if _is_nowrap(shape):
        return

    base_lines, _ = _count_lines(_segments(_para_runs_info(para, original_texts)), box_w)
    base_lines = max(1, base_lines)
    lines, max_word = _count_lines(_segments(_para_runs_info(para)), box_w)
    if lines <= base_lines:
        return

    ratio = 1.0
    while lines > base_lines and ratio > MIN_SHRINK_RATIO + 1e-6:
        ratio = max(MIN_SHRINK_RATIO, ratio * 0.97)
        lines, max_word = _count_lines(_segments(_para_runs_info(para), ratio), box_w)

    if ratio < 1.0:
        for run in runs:
            if run.font.size is not None:
                run.font.size = Pt(round(run.font.size.pt * ratio, 1))
        report.setdefault("ajustes_de_fonte", []).append(
            f"slide {slide_index(slide)} shape {shape.shape_id}: fonte reduzida para {ratio:.0%} "
            f"('{para.text[:40]}')")

    if lines <= base_lines:
        return

    # 2) alargar (texto de 1 linha no template)
    if base_lines == 1:
        needed_pt = sum(w for _, w in _segments(_para_runs_info(para))) / FIT_TOLERANCE
        tf = shape.text_frame
        pad = (tf.margin_left or 91440) + (tf.margin_right or 91440)
        new_w = int(needed_pt * EMU_PER_PT + pad)
        slide_w = slide.part.package.presentation_part.presentation.slide_width
        max_w = int(slide_w * 0.62)
        new_w = min(max(new_w, shape.width), max_w)
        extra = new_w - shape.width
        if extra > 0:
            algn = _para_alignment(para)
            if algn == "ctr":
                shape.left = Emu(int(shape.left - extra / 2))
            elif algn == "r":
                shape.left = Emu(int(shape.left - extra))
            shape.width = Emu(new_w)
            report.setdefault("ajustes_de_largura", []).append(
                f"slide {slide_index(slide)} shape {shape.shape_id}: caixa alargada para caber '{para.text[:40]}'")
            lines, _ = _count_lines(_segments(_para_runs_info(para)), _inner_width_pt(shape) * FIT_TOLERANCE)
            if lines <= base_lines:
                return

    # 3) crescer a altura e empurrar o que está abaixo
    extra_lines = lines - base_lines
    size_pt = max((r.font.size.pt for r in runs if r.font.size is not None), default=18.0)
    extra_h = int(extra_lines * size_pt * LINE_HEIGHT_FACTOR * EMU_PER_PT)
    original_bottom = shape.top + shape.height
    shape.height = Emu(shape.height + extra_h)
    report.setdefault("ajustes_de_altura", []).append(
        f"slide {slide_index(slide)} shape {shape.shape_id}: +{extra_lines} linha(s) para '{para.text[:40]}'")
    _cascade_shift_below(slide, shape, original_bottom, extra_h, report)


def _lock_nowrap(shape):
    """wrap="none" + alinhamento à esquerda no lugar de "justificado" (justificado + quebra manual estica a
    linha, e o LibreOffice recentra caixas sem wrap)."""
    bp = shape.text_frame._txBody.find(f"{{{A_NS}}}bodyPr")
    if bp is not None:
        bp.set("wrap", "none")
        # sem autoajuste: com spAutoFit o LibreOffice encolhe/estica a caixa em torno do CENTRO e o texto "anda"
        for tag in ("spAutoFit", "normAutofit", "noAutofit"):
            for el in bp.findall(f"{{{A_NS}}}{tag}"):
                bp.remove(el)
        bp.insert(0, bp.makeelement(f"{{{A_NS}}}noAutofit", {}))
    for para in shape.text_frame.paragraphs:
        ppr = para._p.find(f"{{{A_NS}}}pPr")
        if ppr is not None and ppr.get("algn") == "just":
            ppr.set("algn", "l")


def fit_paragraph(shape, para, original_texts, slide, report):
    _fit_paragraph_inner(shape, para, original_texts, slide, report)
    # Rede de segurança: texto que no template cabia em UMA linha nunca quebra (nem um dígito "órfão"),
    # em qualquer renderizador. A caixa só cresce na horizontal, para o lado do alinhamento.
    if len(shape.text_frame.paragraphs) == 1 and not _is_nowrap(shape):
        box_w = _inner_width_pt(shape)
        base, _ = _count_lines(_segments(_para_runs_info(para, original_texts)), box_w)
        if base == 1:
            _lock_nowrap(shape)


def slide_index(slide):
    prs = slide.part.package.presentation_part.presentation
    for i, s in enumerate(prs.slides, 1):
        if s.slide_id == slide.slide_id:
            return i
    return 0


def _cascade_shift_below(slide, resized_shape, original_bottom_emu, delta_emu, report=None):
    """Empurra para baixo (mesma coluna, >=30% de sobreposição horizontal) o que estava logo abaixo."""
    if delta_emu <= 0:
        return
    rx1, rx2 = resized_shape.left, resized_shape.left + resized_shape.width
    tolerance = 9525 * 4
    for other in slide.shapes:
        if other.shape_id == resized_shape.shape_id:
            continue
        if other.left is None or other.width is None or other.top is None:
            continue
        # nunca mexe nas laterais decorativas nem em fundos de página inteira
        if other.width >= resized_shape.part.package.presentation_part.presentation.slide_width * 0.9:
            continue
        ox1, ox2 = other.left, other.left + other.width
        overlap = min(rx2, ox2) - max(rx1, ox1)
        min_width = min(resized_shape.width, other.width)
        if min_width <= 0 or overlap / min_width < 0.3:
            continue
        if other.top + tolerance >= original_bottom_emu:
            other.top = Emu(int(other.top) + delta_emu)
            if report is not None:
                report.setdefault("elementos_reposicionados", []).append(
                    f"shape {other.shape_id} deslocado para baixo (crescimento do shape {resized_shape.shape_id})")


# ----------------------------------------------------------------------------
# Valores derivados / consistência (cash-in, percentuais, nº de parcelas)
# ----------------------------------------------------------------------------

def _brl(s):
    """'R$ 1.234.567,89' -> 1234567.89 (None se não der para ler)."""
    if s is None:
        return None
    m = re.search(r"-?\d[\d\.]*(?:,\d+)?", str(s))
    if not m:
        return None
    try:
        return float(m.group(0).replace(".", "").replace(",", "."))
    except ValueError:
        return None


def _fmt_brl(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def derive_fields(data, report):
    """
    Completa/corrige campos que são aritmética pura dos outros, para a proposta
    nunca sair com número do template antigo (ex.: o cash-in do Cristo Rei) nem com
    percentuais que não fecham com os valores.
    """
    d = dict(data)
    notes = report.setdefault("valores_derivados", [])

    # o modelo escreve sempre "+149" (matrículas POTENCIAIS = capacidade livre); o CRM pode mandar sem o sinal
    mp = str(d.get("MATRICULAS_POTENCIAIS", "")).strip()
    if mp and not mp.startswith("+"):
        d["MATRICULAS_POTENCIAIS"] = "+" + mp
        notes.append(f"MATRICULAS_POTENCIAIS padronizado com sinal: '{d['MATRICULAS_POTENCIAIS']}'")

    total = _brl(d.get("VALOR_TOTAL_TRANSACAO"))
    cash_out = _brl(d.get("VALOR_CASH_OUT"))
    if total is not None and cash_out is not None and total > 0:
        cash_in = round(total - cash_out, 2)
        if "VALOR_CASH_IN" not in d or abs((_brl(d.get("VALOR_CASH_IN")) or -1) - cash_in) > 0.01:
            d["VALOR_CASH_IN"] = _fmt_brl(cash_in)
            notes.append(f"VALOR_CASH_IN = total - cash-out = {d['VALOR_CASH_IN']}")
        pct_out = round(cash_out / total * 100)
        pct_in = 100 - pct_out
        for tok, val in (("PERCENTUAL_CASH_OUT", pct_out), ("PERCENTUAL_CASH_IN", pct_in)):
            txt = f"{val}%"
            if d.get(tok) != txt:
                if tok in d:
                    notes.append(f"{tok} recalculado: CRM enviou '{d[tok]}', valores indicam '{txt}'")
                d[tok] = txt

    # entrada total = total destinado - saldo total (quando o CRM não manda)
    for prefix in ("IMOVEL", "ESCOLA"):
        tot = _brl(d.get(f"{prefix}_TOTAL_DESTINADO_VENDEDOR"))
        saldo = _brl(d.get(f"{prefix}_VALOR_SALDO_TOTAL"))
        if f"{prefix}_VALOR_ENTRADA_TOTAL" not in d and tot is not None and saldo is not None:
            d[f"{prefix}_VALOR_ENTRADA_TOTAL"] = _fmt_brl(round(tot - saldo, 2))
            notes.append(f"{prefix}_VALOR_ENTRADA_TOTAL = total destinado - saldo")

    # nº de parcelas (os slides trazem "3 parcelas" / "60 parcelas" fixos no texto)
    for prefix in ("IMOVEL", "ESCOLA"):
        for kind, tot_key, par_key in (("ENTRADA", f"{prefix}_VALOR_ENTRADA_TOTAL", f"{prefix}_VALOR_PARCELA_ENTRADA"),
                                       ("SALDO", f"{prefix}_VALOR_SALDO_TOTAL", f"{prefix}_VALOR_PARCELA_SALDO")):
            t, p = _brl(d.get(tot_key)), _brl(d.get(par_key))
            tok = f"{prefix}_N_PARCELAS_{kind}"
            if tok in d or not t or not p:
                continue
            n = t / p
            if abs(n - round(n)) < 0.02 * max(1, round(n)) and round(n) >= 1:
                d[tok] = str(int(round(n)))
            else:
                report.setdefault("inconsistencias", []).append(
                    f"{tot_key} / {par_key} = {n:.2f} (não é número inteiro de parcelas) — confira os dados do CRM")
    _derive_installment_texts(d, report)
    return d


def _qtd(v):
    m = re.search(r"\d+", str(v)) if v is not None else None
    return int(m.group(0)) if m else None


def _parc(n, valor):
    """'1 parcela de R$ x' / '48 parcelas de R$ x' (singular/plural pela quantidade)."""
    return f"{n} {'parcela' if n == 1 else 'parcelas'} de {valor}"


def _com_pct(v):
    v = str(v).strip()
    return v if v.endswith("%") else v + "%"


def _derive_installment_texts(d, report):
    """
    Quantidades e valores de parcelas dos slides 6 ("Como você recebe") e 7 ("Perspectiva financeira por etapa").
    Nada de número fixo do template: tudo sai dos campos do CRM; singular/plural pela quantidade.
    Nomes antigos do CRM continuam valendo como alternativa (alias).
    """
    alias = {
        "ESCOLA_QTD_PARCELAS_ENTRADA": "ESCOLA_N_PARCELAS_ENTRADA",
        "ESCOLA_QTD_PARCELAS_SALDO": "ESCOLA_N_PARCELAS_SALDO",
        "IMOVEL_QTD_PARCELAS_ENTRADA": "IMOVEL_N_PARCELAS_ENTRADA",
        "IMOVEL_QTD_PARCELAS_SALDO": "IMOVEL_N_PARCELAS_SALDO",
        "MEDIO_PARCELA_ENTRADA_NEGOCIO": "ESCOLA_VALOR_PARCELA_ENTRADA",
        "LONGO_PARCELA_SALDO_NEGOCIO": "ESCOLA_VALOR_PARCELA_SALDO",
        "LONGO_PARCELA_ENTRADA_IMOVEL": "IMOVEL_VALOR_PARCELA_ENTRADA",
        "IMOVEL_VALOR_PARCELA_SALDO": "LONGO_PARCELA_SALDO_IMOVEL",
        "PCT_PARTICIPACAO_MANTENEDOR": "PROLABORE_PARTICIPACAO_PCT",
        "PCT_PARTICIPACAO_ADQUIRIDA": "PROLABORE_LUCRO_PCT",
    }
    for tok, alt in alias.items():
        if not str(d.get(tok, "")).strip() and str(d.get(alt, "")).strip():
            d[tok] = d[alt]
    if not str(d.get("PCT_PARTICIPACAO_ADQUIRIDA", "")).strip() and str(d.get("PERCENTUAL_PARTICIPACAO_SOCIETARIA", "")).strip():
        d["PCT_PARTICIPACAO_ADQUIRIDA"] = d["PERCENTUAL_PARTICIPACAO_SOCIETARIA"]
    for tok in ("PCT_PARTICIPACAO_MANTENEDOR", "PCT_PARTICIPACAO_ADQUIRIDA"):
        if str(d.get(tok, "")).strip():
            d[tok] = _com_pct(d[tok])
    # o slide 6 usa os mesmos números (nome antigo)
    for pre in ("ESCOLA", "IMOVEL"):
        for kind in ("ENTRADA", "SALDO"):
            q = _qtd(d.get(f"{pre}_QTD_PARCELAS_{kind}"))
            if q:
                d[f"{pre}_QTD_PARCELAS_{kind}"] = str(q)
                d[f"{pre}_N_PARCELAS_{kind}"] = str(q)

    def n(k):
        return _qtd(d.get(k))

    def v(k):
        x = str(d.get(k, "")).strip()
        return x or None

    miss = report.setdefault("inconsistencias", [])

    def put(tok, text, needs):
        if text is None:
            miss.append(f"{tok} não gerado: faltam {', '.join(needs)} nos dados")
        else:
            d[tok] = text

    # ---- slide 6: "Como você recebe"
    for pre in ("IMOVEL", "ESCOLA"):
        qe, qs = n(f"{pre}_QTD_PARCELAS_ENTRADA"), n(f"{pre}_QTD_PARCELAS_SALDO")
        if qe:
            d[f"{pre}_TXT_EM_ENTRADA"] = f"em {qe} {'parcela' if qe == 1 else 'parcelas'} de"
            d[f"{pre}_TXT_PRIMEIROS"] = "1 primeiro pagamento" if qe == 1 else f"{qe} primeiros pagamentos"
            d[f"{pre}_TXT_PAGOS_EM"] = "Pago em 1 parcela" if qe == 1 else f"Pagos em {qe} parcelas consecutivas"
        if qs:
            d[f"{pre}_TXT_EM_SALDO"] = f"em {qs} {'parcela' if qs == 1 else 'parcelas'} de"
            d[f"{pre}_TXT_SALDO_MESES"] = f"Saldo em {qs} {'mês' if qs == 1 else 'meses'}"

    # ---- slide 7: perspectiva financeira por etapa
    qe_esc, qs_esc = n("ESCOLA_QTD_PARCELAS_ENTRADA"), n("ESCOLA_QTD_PARCELAS_SALDO")
    qe_imo, qs_imo = n("IMOVEL_QTD_PARCELAS_ENTRADA"), n("IMOVEL_QTD_PARCELAS_SALDO")
    v_ent_esc, v_sal_esc = v("MEDIO_PARCELA_ENTRADA_NEGOCIO"), v("LONGO_PARCELA_SALDO_NEGOCIO")
    v_ent_imo, v_sal_imo = v("LONGO_PARCELA_ENTRADA_IMOVEL"), v("IMOVEL_VALOR_PARCELA_SALDO")

    def ini(q):
        return "a parcela inicial" if q == 1 else f"as {q} parcelas iniciais"

    def titulo(q):
        return "Total mensal na parcela inicial" if q == 1 else f"Total mensal nas {q} parcelas iniciais"

    put("S7_MEDIO_ENTRADA_NEGOCIO", _parc(qe_esc, v_ent_esc) if qe_esc and v_ent_esc else None,
        ["ESCOLA_QTD_PARCELAS_ENTRADA", "MEDIO_PARCELA_ENTRADA_NEGOCIO"])
    put("S7_LONGO_PARCELAS_NEGOCIO", _parc(qs_esc, v_sal_esc) if qs_esc and v_sal_esc else None,
        ["ESCOLA_QTD_PARCELAS_SALDO", "LONGO_PARCELA_SALDO_NEGOCIO"])
    put("S7_LONGO_ENTRADA_IMOVEL", _parc(qe_imo, v_ent_imo) if qe_imo and v_ent_imo else None,
        ["IMOVEL_QTD_PARCELAS_ENTRADA", "LONGO_PARCELA_ENTRADA_IMOVEL"])
    put("S7_NOTA_MEDIO",
        f"Após {ini(qe_esc)}, o saldo do negócio passa para {_parc(qs_esc, v_sal_esc)}."
        if qe_esc and qs_esc and v_sal_esc else None,
        ["ESCOLA_QTD_PARCELAS_ENTRADA", "ESCOLA_QTD_PARCELAS_SALDO", "LONGO_PARCELA_SALDO_NEGOCIO"])
    put("S7_NOTA_LONGO",
        f"Após {ini(qe_imo)}, os saldos passam para {_parc(qs_esc, v_sal_esc)} no negócio e "
        f"{_parc(qs_imo, v_sal_imo)} no imóvel."
        if qe_imo and qs_esc and v_sal_esc and qs_imo and v_sal_imo else None,
        ["IMOVEL_QTD_PARCELAS_ENTRADA", "ESCOLA_QTD_PARCELAS_SALDO", "LONGO_PARCELA_SALDO_NEGOCIO",
         "IMOVEL_QTD_PARCELAS_SALDO", "IMOVEL_VALOR_PARCELA_SALDO"])
    put("S7_TITULO_TOTAL_MEDIO", titulo(qe_esc) if qe_esc else None, ["ESCOLA_QTD_PARCELAS_ENTRADA"])
    put("S7_TITULO_TOTAL_LONGO", titulo(qe_imo) if qe_imo else None, ["IMOVEL_QTD_PARCELAS_ENTRADA"])


# ----------------------------------------------------------------------------
# Aplicação do field_map
# ----------------------------------------------------------------------------

def apply_field_map(prs, field_map, data, report):
    """Percorre cada regra do field_map e substitui o texto no shape exato."""
    slides = list(prs.slides)
    originals = {}   # id(p) -> (shape, para, slide, [textos originais dos runs])

    for rule in field_map:
        token = rule["token"]
        slide_no = rule["slide"]
        shape_id = rule["shape_id"]
        mode = rule["mode"]
        find = rule["find"]

        if token not in data:
            if not rule.get("optional"):
                report["missing_tokens"].add(token)
            continue
        value = str(data[token])
        if rule.get("fmt"):
            value = rule["fmt"].format(v=value)

        if slide_no < 1 or slide_no > len(slides):
            report["errors"].append(f"slide {slide_no} não existe (regra do token {token})")
            continue
        slide = slides[slide_no - 1]
        shape = _find_shape_by_id(slide.shapes, shape_id)
        if shape is None:
            report["errors"].append(f"shape_id {shape_id} não encontrado no slide {slide_no} (token {token})")
            continue
        if not shape.has_text_frame:
            report["errors"].append(f"shape_id {shape_id} no slide {slide_no} não é uma caixa de texto (token {token})")
            continue

        applied = False
        for para in shape.text_frame.paragraphs:
            key = id(para._p)
            if key not in originals:
                originals[key] = (shape, para, slide, [r.text for r in para.runs])
            para_hit = False
            for run in para.runs:
                if mode == "exact" and run.text == find:
                    run.text = value
                    applied = para_hit = True
                elif mode == "contains" and find in run.text:
                    run.text = _replace_keeping_context(run.text, find, value)
                    applied = para_hit = True
            # o Canva às vezes parte uma frase em vários runs (ex.: "...do Colégio" + " Cristo Rei, ...")
            if mode == "contains" and not para_hit and find in para.text:
                if _replace_across_runs(para, find, value):
                    applied = True
        if applied:
            report["applied"] += 1
        elif not rule.get("optional"):
            report["not_found"].append(
                f"token {token}: texto '{find}' (mode={mode}) não encontrado no shape_id {shape_id} "
                f"do slide {slide_no} — o texto do template pode ter mudado. Verifique field_map.json.")

    # UMA passada de ajuste por parágrafo, depois de todas as trocas
    touched = report.setdefault("_touched", set())
    for shape, para, slide, orig in originals.values():
        if [r.text for r in para.runs] != orig:
            touched.add((slide_index(slide), shape.shape_id))
            fit_paragraph(shape, para, orig, slide, report)


def _replace_across_runs(para, find, value):
    runs = list(para.runs)
    full = "".join(r.text for r in runs)
    i = full.find(find)
    if i < 0:
        return False
    j = i + len(find)
    pos, first = 0, True
    for r in runs:
        start, end = pos, pos + len(r.text)
        pos = end
        if end <= i or start >= j:
            continue
        a, b = max(i, start) - start, min(j, end) - start
        r.text = r.text[:a] + (value if first else "") + r.text[b:]
        first = False
    return True


def _replace_keeping_context(text, find, value):
    """
    Troca `find` por `value` sem duplicar o que já está ao redor. Ex.: o template tem
    '3 parcelas de R$ 20.000,00'; se o CRM mandar 'R$ 30.000,00' ficamos com
    '3 parcelas de R$ 30.000,00'; se mandar '3 parcelas de R$ 30.000,00' também.
    """
    i = text.index(find)
    before, after = text[:i], text[i + len(find):]
    b = before.strip()
    if b and value.lstrip().startswith(b):
        value = value.lstrip()[len(b):].lstrip()
    a = after.strip()
    if a and value.rstrip().endswith(a):
        value = value.rstrip()[: -len(a)].rstrip()
    return before + value + after


def _find_shape_by_id(shapes, target_id):
    for shape in shapes:
        if shape.shape_id == target_id:
            return shape
        if shape.shape_type == 6:
            found = _find_shape_by_id(shape.shapes, target_id)
            if found is not None:
                return found
    return None


# ----------------------------------------------------------------------------
# Laterais decorativas: UMA imagem por lado (idêntica ao Canva em qualquer renderizador)
# ----------------------------------------------------------------------------

def replace_edges(prs, report, edges_dir=EDGES_DIR):
    done = 0
    for idx, slide in enumerate(prs.slides, 1):
        left_png, right_png = Path(edges_dir) / f"s{idx}_L.png", Path(edges_dir) / f"s{idx}_R.png"
        if not (left_png.exists() and right_png.exists()):
            continue
        tree = slide.shapes._spTree
        victims = []
        for sh in slide.shapes:
            if sh.left is None or sh.width is None:
                continue
            if sh.has_text_frame and sh.text_frame.text.strip():
                continue
            if sh.left + sh.width <= EDGE_LEFT_LIMIT_EMU or sh.left >= EDGE_RIGHT_LIMIT_EMU:
                victims.append(sh)
        if not victims:
            continue
        insert_at = min(list(tree).index(v._element) for v in victims)
        for v in victims:
            tree.remove(v._element)
        slide_h = prs.slide_height
        slide_w = prs.slide_width
        for png, left in ((left_png, 0), (right_png, None)):
            with Image.open(png) as im:
                wpx, hpx = im.size
            emu_px = slide_h / hpx
            w_emu = int(round(wpx * emu_px))
            x = 0 if left == 0 else slide_w - w_emu
            pic = slide.shapes.add_picture(str(png), Emu(x), Emu(0), Emu(w_emu), Emu(slide_h))
            pic.name = "Lateral decorativa"
            tree.remove(pic._element)
            tree.insert(insert_at, pic._element)
            insert_at += 1
        done += 1
    # capa e fecho: peça laranja do canto direito
    for idx in (1, 10):
        png = Path(edges_dir) / f"s{idx}_corner.png"
        if not png.exists() or idx > len(prs.slides):
            continue
        slide = list(prs.slides)[idx - 1]
        for sh in list(slide.shapes):
            if sh.left is None or (sh.has_text_frame and sh.text_frame.text.strip()):
                continue
            if sh.left == 17732542 and sh.top == 2205962:
                tree = slide.shapes._spTree
                pos = list(tree).index(sh._element)
                tree.remove(sh._element)
                with Image.open(png) as im:
                    wpx, hpx = im.size
                w_emu = prs.slide_width - 17732542
                pic = slide.shapes.add_picture(str(png), Emu(17732542), Emu(2205962), Emu(w_emu), Emu(int(w_emu * hpx / wpx)))
                pic.name = "Canto decorativo"
                tree.remove(pic._element)
                tree.insert(pos, pic._element)
                done += 1
                break
    report["laterais_unificadas_em_slides"] = done


# ----------------------------------------------------------------------------
# Foto da fachada: recorte centralizado para o enquadramento de CADA slide
# ----------------------------------------------------------------------------

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _iter_shapes_abs(shapes, tx=(0.0, 0.0, 1.0, 1.0)):
    """
    Percorre shapes (inclusive dentro de grupos) devolvendo (shape, left, top, width, height) em
    coordenadas ABSOLUTAS do slide. tx = (offset_x, offset_y, scale_x, scale_y) do grupo-pai.
    """
    ox, oy, sx, sy = tx
    for sh in shapes:
        if sh.left is None:
            continue
        left, top = ox + sh.left * sx, oy + sh.top * sy
        w, h = sh.width * sx, sh.height * sy
        yield sh, left, top, w, h
        if sh.shape_type == 6:          # GROUP: converte do espaço-filho para o espaço do slide
            xfrm = sh._element.grpSpPr.find(f"{{{A_NS}}}xfrm")
            ch_off, ch_ext = xfrm.find(f"{{{A_NS}}}chOff"), xfrm.find(f"{{{A_NS}}}chExt")
            csx = w / int(ch_ext.get("cx"))
            csy = h / int(ch_ext.get("cy"))
            ntx = (left - int(ch_off.get("x")) * csx, top - int(ch_off.get("y")) * csy, csx, csy)
            yield from _iter_shapes_abs(sh.shapes, ntx)


def replace_photo(prs, photo_path, report, media_filename="image14.png"):
    """
    A foto da fachada aparece nos slides 2, 3 e 9 como PREENCHIMENTO de uma forma livre (dentro de um
    grupo que sai parcialmente da lâmina), com recorte definido por fillRect. Se só trocarmos o arquivo
    de mídia, a foto nova herda o recorte da foto do Cristo Rei e sai esticada/cortada.
    Aqui: (1) calculamos a parte VISÍVEL da moldura em cada slide, (2) recortamos a foto nova (centro,
    levemente para cima) exatamente nessa proporção e (3) fazemos o preenchimento ocupar só a parte visível.
    """
    if not photo_path:
        report["photo"] = "AVISO: nenhuma foto enviada — a foto do template (Cristo Rei) foi mantida"
        return
    with Image.open(photo_path) as im:
        src = im.convert("RGB")
    sw, sh_ = prs.slide_width, prs.slide_height
    used = 0
    for slide in prs.slides:
        for shape, left, top, w, h in _iter_shapes_abs(slide.shapes):
            sppr = shape._element.find(f"{{http://schemas.openxmlformats.org/presentationml/2006/main}}spPr")
            if sppr is None:
                continue
            bf = sppr.find(f"{{{A_NS}}}blipFill")
            blip = bf.find(f"{{{A_NS}}}blip") if bf is not None else None
            if blip is None:
                continue
            rid = blip.get(f"{{{R_NS}}}embed")
            if not rid or not str(slide.part.related_part(rid).partname).endswith(media_filename):
                continue
            # parte visível da moldura (fração em cada lado)
            vl, vt = max(0.0, left), max(0.0, top)
            vr, vb = min(float(sw), left + w), min(float(sh_), top + h)
            vis_w, vis_h = vr - vl, vb - vt
            if vis_w <= 0 or vis_h <= 0:
                continue
            fl, ft = (vl - left) / w, (vt - top) / h
            fr, fb = (left + w - vr) / w, (top + h - vb) / h

            target = vis_w / vis_h
            iw, ih = src.size
            if iw / ih > target:
                nw = int(ih * target); x0 = (iw - nw) // 2; box = (x0, 0, x0 + nw, ih)
            else:
                nh = int(iw / target); y0 = int((ih - nh) * 0.35); box = (0, y0, iw, y0 + nh)
            crop = src.crop(box)
            vw = min(2400, crop.width)
            vh = int(vw / target)
            crop = crop.resize((vw, vh), Image.LANCZOS)
            # A moldura sai parcialmente da lâmina (parte de cima/lados fica fora do slide). Montamos uma imagem
            # do tamanho TOTAL da moldura, com a foto na parte visível e um "preenchimento" desfocado na parte
            # invisível, e usamos fillRect=0 (sem recorte) — o LibreOffice não trata bem fillRect positivo.
            full_w = int(round(vw / max(1e-6, (1 - fl - fr))))
            full_h = int(round(vh / max(1e-6, (1 - ft - fb))))
            canvas = crop.resize((full_w, full_h)).filter(ImageFilter.GaussianBlur(25))
            canvas.paste(crop, (int(round(fl * full_w)), int(round(ft * full_h))))
            buf = BytesIO(); canvas.save(buf, format="JPEG", quality=93); buf.seek(0)
            _, new_rid = slide.part.get_or_add_image_part(buf)
            blip.set(f"{{{R_NS}}}embed", new_rid)
            for child in list(blip):
                blip.remove(child)
            stretch = bf.find(f"{{{A_NS}}}stretch")
            if stretch is None:
                stretch = bf.makeelement(f"{{{A_NS}}}stretch", {}); bf.append(stretch)
            for old in stretch.findall(f"{{{A_NS}}}fillRect"):
                stretch.remove(old)
            stretch.append(stretch.makeelement(f"{{{A_NS}}}fillRect", {}))
            used += 1
    report["photo"] = f"foto aplicada em {used} moldura(s) com recorte centralizado"
    if used == 0:
        report["errors"].append(f"Nenhuma moldura usa {media_filename} — a foto NÃO foi trocada. Confira o template.")



def normalize_line_breaks(prs, field_map=None, pills_path=PILLS_JSON, breaks_path=BREAKS_JSON):
    """
    O .pptx exportado pelo Canva não traz as quebras de linha que o Canva usou no design; cada renderizador
    (PowerPoint, LibreOffice...) quebra num ponto um pouco diferente. Gravamos as quebras do PDF do Canva como
    <a:br/> e travamos o wrap — o texto quebra IGUAL em qualquer lugar. Caixas cujo texto muda com os dados
    (field_map) só recebem as quebras se tiverem fundo de texto (pills.json), pois o fundo precisa casar com as linhas.
    """
    slides = list(prs.slides)
    dynamic = {(r["slide"], r["shape_id"]) for r in (field_map or [])}
    # caixas cujo texto muda de ESTRUTURA (ex.: 'Pagos em 3 parcelas consecutivas' -> 'Pago em 1 parcela'):
    # as quebras do Canva não valem; o texto quebra natural e o fundo é recalculado pela medição real
    free = {(r["slide"], r["shape_id"]) for r in (field_map or []) if r.get("free_wrap")}
    work = {}
    if Path(breaks_path).exists():
        for sno, shapes in json.load(open(breaks_path, encoding="utf-8")).items():
            for sid, lines in shapes.items():
                if (int(sno), int(sid)) not in dynamic:
                    work[(int(sno), int(sid))] = lines
    if Path(pills_path).exists():
        for pill in json.load(open(pills_path, encoding="utf-8")):
            if len(pill["lines"]) >= 2 and (pill["slide"], pill["shape_id"]) not in free:
                work[(pill["slide"], pill["shape_id"])] = [ln["text"] for ln in pill["lines"]]
    done = 0
    for (sno, sid), line_texts in work.items():
        shape = _find_shape_by_id(slides[sno - 1].shapes, sid)
        if shape is None or not shape.has_text_frame:
            continue
        if _insert_breaks(shape, line_texts):
            _lock_nowrap(shape)
            done += 1
    return done


def _insert_breaks(shape, line_texts):
    targets = [re.sub(r"\s+", "", t) for t in line_texts]
    li, plan = 0, []
    for para in shape.text_frame.paragraphs:
        text = "".join(r.text for r in para.runs)
        if not text.strip():
            continue
        offsets, pos = [], 0
        while li < len(targets):
            need, consumed, j = len(targets[li]), 0, pos
            while j < len(text) and consumed < need:
                if not text[j].isspace():
                    if text[j] != targets[li][consumed]:
                        return False
                    consumed += 1
                j += 1
            if consumed < need:
                return False
            li += 1
            if j >= len(text.rstrip()):
                break
            offsets.append(j)
            while j < len(text) and text[j] == " ":
                j += 1
            pos = j
        plan.append((para, offsets))
    if li != len(targets):
        return False
    for para, offsets in plan:
        for off in reversed(offsets):
            _split_run_with_break(para, off)
    return True


def finalize_nowrap_boxes(prs):
    """
    Caixas com wrap="none" são redimensionadas pelo LibreOffice em torno do CENTRO: se a caixa é mais larga
    que o texto, o texto "anda" para o lado. Ajustamos a largura da caixa ao texto (mantendo a âncora do
    alinhamento) para que ele fique exatamente onde o Canva o colocou.
    """
    for slide in prs.slides:
        for sh, *_ in _iter_shapes_abs(slide.shapes):
            if not (sh.has_text_frame and sh.text_frame.text.strip()) or not _is_nowrap(sh) or sh.shape_type == 6:
                continue
            widths = _wrapped_line_widths(sh)
            if not widths:
                continue
            tf = sh.text_frame
            pad = ((tf.margin_left if tf.margin_left is not None else 91440) + (tf.margin_right if tf.margin_right is not None else 91440))
            new_w = int(max(widths) * 1.02 * EMU_PER_PT + 2 * EMU_PER_PT + pad)
            old_w = sh.width
            if abs(new_w - old_w) < 3 * EMU_PER_PT:
                continue
            algn = _para_alignment(tf.paragraphs[0])
            if algn == "ctr":
                sh.left = Emu(int(sh.left + (old_w - new_w) / 2))
            elif algn == "r":
                sh.left = Emu(int(sh.left + (old_w - new_w)))
            sh.width = Emu(new_w)


def _split_run_with_break(para, offset):
    """Insere <a:br/> na posição `offset` (em caracteres) do texto do parágrafo, removendo o espaço da quebra."""
    import copy
    pos = 0
    for r in para._p.findall(f"{{{A_NS}}}r"):
        t = r.find(f"{{{A_NS}}}t")
        txt = t.text or ""
        if pos <= offset <= pos + len(txt):
            k = offset - pos
            left, right = txt[:k].rstrip(" "), txt[k:].lstrip(" ")
            br = r.makeelement(f"{{{A_NS}}}br", {})
            rpr = r.find(f"{{{A_NS}}}rPr")
            if rpr is not None:
                br.append(copy.deepcopy(rpr))
            if right:
                r2 = copy.deepcopy(r)
                r2.find(f"{{{A_NS}}}t").text = right
                r.addnext(r2)
            r.addnext(br)
            t.text = left
            if not left:
                para._p.remove(r)
            return
        pos += len(txt)


# ----------------------------------------------------------------------------
# "Marca-textos" (fundos arredondados atrás do texto) que o Canva não exporta no .pptx
# ----------------------------------------------------------------------------

def _wrapped_line_widths(shape):
    """Largura (pt) de cada linha visual do texto atual da caixa, com as fontes reais."""
    nowrap = _is_nowrap(shape)
    box = 1e9 if nowrap else _inner_width_pt(shape)
    widths = []
    for para in shape.text_frame.paragraphs:
        words = _segments(_para_runs_info(para))
        cur, pend, any_word = 0.0, 0.0, False
        for text, w in words:
            if text == "\n":
                widths.append(cur)
                cur, pend, any_word = 0.0, 0.0, False
                continue
            if text == " ":
                pend += w
                continue
            any_word = True
            if cur > 0 and cur + pend + w > box:
                widths.append(cur)
                cur = w
            else:
                cur += (pend if cur > 0 else 0) + w
            pend = 0.0
        if any_word:
            widths.append(cur)
    return widths


def _top_level_element(shape):
    el = shape._element
    while el.getparent() is not None and el.getparent().tag.split("}")[1] != "spTree":
        el = el.getparent()
    return el


def add_pills(prs, report, pills_path=PILLS_JSON):
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    if not Path(pills_path).exists():
        return
    pills = json.load(open(pills_path, encoding="utf-8"))
    slides = list(prs.slides)
    touched = report.get("_touched", set())
    count = 0
    for pill in pills:
        slide = slides[pill["slide"] - 1]
        found = None
        for sh, l, t, w, h in _iter_shapes_abs(slide.shapes):
            if sh.shape_id == pill["shape_id"]:
                found = (sh, l / EMU_PER_PT, t / EMU_PER_PT, w / EMU_PER_PT)
                break
        if not found:
            continue
        shape, left_pt, top_pt, width_pt = found
        dx, dy = left_pt - pill["shape_left"], top_pt - pill["shape_top"]
        stored = pill["lines"]
        is_touched = (pill["slide"], pill["shape_id"]) in touched

        # escala (se a fonte foi reduzida no ajuste automático)
        scale = 1.0
        for run in (r for pa in shape.text_frame.paragraphs for r in pa.runs):
            if run.font.size is not None and pill["size_pt"]:
                scale = min(1.0, run.font.size.pt / pill["size_pt"])
                break
        padx, pad_r = pill["padx"] * scale, pill["pad_r"] * scale

        widths = _wrapped_line_widths(shape) if is_touched else None
        n = len(widths) if widths else len(stored)
        pitch = stored[-1]["bottom"] - stored[-1]["top"]
        for i in range(n):
            ref = stored[min(i, len(stored) - 1)]
            top = ref["top"] + dy + max(0, i - (len(stored) - 1)) * pitch
            bottom = top + (ref["bottom"] - ref["top"])
            if widths is None:
                x0, x1 = ref["x0"] + dx - pill["padx"], ref["x1"] + dx + pill["pad_r"]
            else:
                tw = widths[i]
                if pill["algn"] == "ctr":
                    cx = left_pt + width_pt / 2
                    x0, x1 = cx - tw / 2 - padx, cx + tw / 2 + pad_r
                elif pill["algn"] == "r":
                    xr = left_pt + width_pt
                    x0, x1 = xr - tw - padx, xr + pad_r
                else:
                    tx0 = stored[0]["x0"] + dx
                    x0, x1 = tx0 - padx, tx0 + tw + pad_r
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(int(x0 * EMU_PER_PT)), Emu(int(top * EMU_PER_PT)),
                                         Emu(int((x1 - x0) * EMU_PER_PT)), Emu(int((bottom - top) * EMU_PER_PT)))
            box.name = "Fundo do texto"
            box.adjustments[0] = 0.27
            box.fill.solid(); box.fill.fore_color.rgb = RGBColor.from_string(pill["color"])
            box.line.fill.background()
            style = box._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}style')
            if style is not None:
                box._element.remove(style)
            tree = slide.shapes._spTree
            tree.remove(box._element)
            target = _top_level_element(shape)
            tree.insert(list(tree).index(target), box._element)
            count += 1
    report["fundos_de_texto"] = count


# ----------------------------------------------------------------------------
# Slide 2: grade de 87 bonequinhos proporcional à ocupação + colchetes + rótulos
# ----------------------------------------------------------------------------
ICON_W, ICON_H = 204157, 281596          # tamanho de cada bonequinho (EMU)
ICON_ID_RANGE = (32, 119)                # ids dos bonequinhos no slide 2 (68 = colchete laranja)
BRACKETS = {                             # medidas extraídas do PDF do Canva (pt)
    "orange": dict(shape_id=68, color="FF5B22", dot=5.08, thick=1.52, gap_right=10.32, left_frac=0.449,
                   y_line=697.46, y_top=664.91, label_id=133),
    "grey": dict(shape_id=120, color="9A9C9F", dot=6.0, thick=1.80, gap_right=11.76, left_frac=0.543,
                 y_line=696.99, y_top=658.54, label_id=136),
}
LABEL_MARGIN_EMU = 1482747               # margem esquerda do conteúdo


def _pct(text):
    m = re.search(r"\d+(?:[\.,]\d+)?", str(text or ""))
    if not m:
        return None
    try:
        return float(m.group(0).replace(".", "").replace(",", ".")) if "," in m.group(0) else float(m.group(0))
    except ValueError:
        return None


def apply_capacity_graphic(prs, data, report):
    """
    O template desenha 87 bonequinhos (29 colunas x 3): 36 laranja ("alunos atuais") + 51 cinza ("matrículas
    potenciais") = 41,4%, exatamente a ocupação do Cristo Rei. Aqui a divisão laranja/cinza passa a seguir a
    ocupação da escola; o espaço entre os dois blocos, os colchetes e os rótulos acompanham.
    """
    import copy
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    pct = _pct(data.get("PERCENTUAL_OCUPACAO"))
    if pct is None or len(prs.slides) < 2:
        return
    slide = list(prs.slides)[1]
    by_id = {sh.shape_id: sh for sh in slide.shapes}
    icons = [sh for i, sh in by_id.items() if ICON_ID_RANGE[0] <= i <= ICON_ID_RANGE[1] and i != 68
             and sh.width == ICON_W and sh.height == ICON_H]
    if len(icons) != 87 or 32 not in by_id:
        report.setdefault("avisos_gerais", []).append("Grade de bonequinhos do slide 2 não reconhecida — mantida como no template.")
        return
    # colunas (agrupa x próximos) e ordem coluna-a-coluna, de cima para baixo
    xs = sorted({sh.left for sh in icons})
    groups, cur = [], [xs[0]]
    for x in xs[1:]:
        if x - cur[-1] < 100000:
            cur.append(x)
        else:
            groups.append(cur); cur = [x]
    groups.append(cur)
    if len(groups) != 29:
        return
    col_of = {x: i for i, g in enumerate(groups) for x in g}
    col_x = [min(g) for g in groups]
    pitch = (col_x[11] - col_x[0]) / 11
    gap_model = col_x[12] - col_x[11] - pitch
    base = [col_x[i] if i < 12 else col_x[i] - gap_model for i in range(29)]

    n_orange = max(0, min(87, int(round(pct / 100 * 87))))
    k_orange = -(-n_orange // 3)                      # colunas que têm laranja
    shift = gap_model if 0 < k_orange < 29 else 0
    orange_blip = copy.deepcopy(by_id[32]._element.spPr.find(f"{{{A_NS}}}blipFill").find(f"{{{A_NS}}}blip"))
    grey_src = next(sh for sh in icons if sh.left >= col_x[12] - 1)
    grey_blip = copy.deepcopy(grey_src._element.spPr.find(f"{{{A_NS}}}blipFill").find(f"{{{A_NS}}}blip"))

    ordered = sorted(icons, key=lambda sh: (col_of[min(g for g in groups[col_of[sh.left]])] if False else col_of[sh.left], sh.top))
    for idx, sh in enumerate(ordered):
        c = col_of[sh.left]
        sh.left = Emu(int(base[c] + (shift if c >= k_orange else 0)))
        bf = sh._element.spPr.find(f"{{{A_NS}}}blipFill")
        old = bf.find(f"{{{A_NS}}}blip")
        bf.replace(old, copy.deepcopy(orange_blip if idx < n_orange else grey_blip))

    # blocos
    def extent(cols):
        if not cols:
            return None
        xs_ = [base[c] + (shift if c >= k_orange else 0) for c in cols]
        return min(xs_), max(xs_) + ICON_W
    blocks = {"orange": extent(range(0, k_orange)), "grey": extent(range(k_orange, 29))}

    tree = slide.shapes._spTree
    for name, cfg in BRACKETS.items():
        old = by_id.get(cfg["shape_id"])
        pos = list(tree).index(old._element) if old is not None else len(tree)
        if old is not None:
            tree.remove(old._element)
        blk = blocks[name]
        label = by_id.get(cfg["label_id"])
        if blk is None:
            continue
        left_pt, right_pt = blk[0] / EMU_PER_PT, blk[1] / EMU_PER_PT
        w_pt = right_pt - left_pt
        x_r = right_pt + cfg["gap_right"]
        x_l = min(left_pt + cfg["left_frac"] * w_pt, x_r - 18)
        t, d = cfg["thick"], cfg["dot"]

        def box(kind, x, y, w, h):
            shp = slide.shapes.add_shape(kind, Emu(int(x * EMU_PER_PT)), Emu(int(y * EMU_PER_PT)),
                                         Emu(int(w * EMU_PER_PT)), Emu(int(h * EMU_PER_PT)))
            shp.fill.solid(); shp.fill.fore_color.rgb = RGBColor.from_string(cfg["color"])
            shp.line.fill.background()
            style = shp._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}style")
            if style is not None:
                shp._element.remove(style)
            shp.name = f"Colchete {name}"
            tree.remove(shp._element)
            return shp._element
        parts = [
            box(MSO_SHAPE.RECTANGLE, x_l, cfg["y_line"] - t / 2, x_r - x_l, t),
            box(MSO_SHAPE.RECTANGLE, x_r - t / 2, cfg["y_top"], t, cfg["y_line"] - cfg["y_top"] + t / 2),
            box(MSO_SHAPE.OVAL, x_l - d / 2, cfg["y_line"] - d / 2, d, d),
            box(MSO_SHAPE.OVAL, x_r - d / 2, cfg["y_top"] - d / 2, d, d),
        ]
        for k, el in enumerate(parts):
            tree.insert(pos + k, el)
        # rótulo centralizado sob o bloco (sem passar da margem esquerda)
        if label is not None and label.has_text_frame:
            widths = _wrapped_line_widths(label)
            tw = max(widths) if widths else label.width / EMU_PER_PT
            center = (blk[0] + blk[1]) / 2
            new_w = int((tw * 1.02 + 2) * EMU_PER_PT)
            label.width = Emu(new_w)
            label.left = Emu(int(max(LABEL_MARGIN_EMU, center - new_w / 2)))
    report["grade_bonequinhos"] = f"{n_orange} laranja / {87 - n_orange} cinza ({pct:.2f}% de ocupação)"


# ----------------------------------------------------------------------------
# Pipeline completo
# ----------------------------------------------------------------------------

def protect_static_slides(prs):
    """Textos soltos do logo ("Educação") e da capa não podem quebrar linha por diferença de métrica de fonte."""
    for idx, slide in enumerate(prs.slides, 1):
        for sh, *_ in _iter_shapes_abs(slide.shapes):
            if not (sh.has_text_frame and sh.text_frame.text.strip()):
                continue
            if idx == 1 or sh.text_frame.text.strip() == "Educação":
                _lock_nowrap(sh)


def strip_synthetic_bold(prs):
    """
    Todas as fontes do template já são a variante pesada no próprio nome ("Eastman Alternate Trial Bold",
    "Neulis Neue Bold"...) e o Canva ainda marca b="true". Sem a fonte em negrito de verdade, o LibreOffice
    e o PowerPoint "engrossam" o desenho por cima (negrito sintético) e o texto fica mais pesado/largo que
    no Canva. Tirando o b="true" dessas fontes o texto sai com o peso e a largura do design original.
    """
    n = 0
    for slide in prs.slides:
        for rpr in slide._element.iter(f"{{{A_NS}}}rPr", f"{{{A_NS}}}endParaRPr", f"{{{A_NS}}}defRPr"):
            lat = rpr.find(f"{{{A_NS}}}latin")
            face = lat.get("typeface") if lat is not None else ""
            if re.search(r"Bold|Semi|Medium", face or "") and rpr.get("b") in ("1", "true"):
                rpr.set("b", "0")
                n += 1
    return n


def apply_line_pitch(prs, pitch_path=PITCH_JSON):
    """Usa o espaçamento entre linhas medido no PDF do Canva (ver tools/build_pitch.py)."""
    if not Path(pitch_path).exists():
        return 0
    table, n = json.load(open(pitch_path)), 0
    slides = list(prs.slides)
    for sno, shapes in table.items():
        for sid, pitch in shapes.items():
            shape = _find_shape_by_id(slides[int(sno) - 1].shapes, int(sid))
            if shape is None or not shape.has_text_frame:
                continue
            for el in shape.text_frame._txBody.iter(f"{{{A_NS}}}lnSpc"):
                pts = el.find(f"{{{A_NS}}}spcPts")
                if pts is not None:
                    pts.set("val", str(int(round(pitch * 100))))
                    n += 1
    return n


def _flip_rotated_pictures(prs):
    import io
    from PIL import Image
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    done = set()
    for slide in prs.slides:
        for xfrm in slide._element.iter(A + "xfrm"):
            rot = xfrm.get("rot")
            if not (rot and rot.lstrip("-").isdigit() and int(rot) % 21600000 == 10800000):
                continue
            geom = xfrm.getnext()
            blip = None
            if geom is not None:
                blip = geom.getnext().find(A + "blip") if geom.getnext() is not None and geom.getnext().tag == A + "blipFill" else None
            if blip is None:
                continue
            rid = blip.get(R + "embed")
            part = slide.part.related_part(rid)
            key = id(part)
            if key not in done:
                try:
                    img = Image.open(io.BytesIO(part.blob)).convert("RGBA").rotate(180)
                    buf = io.BytesIO()
                    img.save(buf, "PNG")
                    part._blob = buf.getvalue()
                except Exception:
                    continue
                done.add(key)
            for ext in blip.findall(A + "extLst"):
                blip.remove(ext)
            xfrm.set("rot", "0")


def tune_for_libreoffice(src_pptx, dst_pptx, offsets_path=LO_OFFSETS_JSON):
    """
    Cria uma cópia do .pptx pronta para o LibreOffice gerar o PDF: sobe/desce cada caixa de texto pelo
    deslocamento medido contra o PDF do Canva (assets/lo_offsets.json). O .pptx entregue ao cliente NÃO
    recebe isso (no PowerPoint o texto já cai no lugar certo).
    """
    prs = Presentation(str(src_pptx))
    if Path(offsets_path).exists():
        table = json.load(open(offsets_path))
        slides = list(prs.slides)
        for sno, shapes in table.items():
            slide = slides[int(sno) - 1]
            by_id = {}
            for sh, l, t, w, h in _iter_shapes_abs(slide.shapes):
                by_id[sh.shape_id] = (sh, (h / sh.height) if sh.height else 1.0)
            for sid, dy in shapes.items():
                if int(sid) not in by_id:
                    continue
                sh, scale = by_id[int(sid)]
                if sh.top is None:
                    continue
                sh.top = Emu(int(sh.top - dy * EMU_PER_PT / (scale or 1.0)))
    # o LibreOffice desenha o preenchimento por imagem sem a rotação de 180° (rot="-10800000"): as setas
    # tracejadas "Curto → Médio → Longo" saíam viradas para trás no PDF. Na cópia do PDF a imagem é girada de
    # verdade (e o SVG, que o LibreOffice preferiria, é descartado) e a rotação da forma vira 0.
    _flip_rotated_pictures(prs)
    prs.save(str(dst_pptx))


def new_report():
    return {"applied": 0, "not_found": [], "missing_tokens": set(), "errors": [], "photo": None}


def generate(template_path, field_map, data, out_path, photo_path=None, report=None):
    report = report if report is not None else new_report()
    prs = Presentation(str(template_path))
    data = derive_fields(data, report)
    normalize_line_breaks(prs, field_map)
    apply_line_pitch(prs)
    apply_field_map(prs, field_map, data, report)
    finalize_nowrap_boxes(prs)
    apply_capacity_graphic(prs, data, report)
    add_pills(prs, report)
    replace_edges(prs, report)
    protect_static_slides(prs)
    strip_synthetic_bold(prs)
    replace_photo(prs, photo_path, report)
    prs.save(str(out_path))
    return report


def main():
    ap = argparse.ArgumentParser(description="Mail-merge para o template Rede Lumo (Canva)")
    ap.add_argument("--template", required=True)
    ap.add_argument("--data", required=True)
    ap.add_argument("--map", default="field_map.json")
    ap.add_argument("--foto", default=None)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    report = generate(args.template, load_json(args.map), load_json(args.data), args.out, args.foto)

    print(f"\n✅ {report['applied']} campos substituídos.")
    for key, title in (("missing_tokens", "Tokens ausentes"), ("not_found", "Regras sem correspondência"),
                       ("errors", "Erros"), ("valores_derivados", "Valores derivados/corrigidos"),
                       ("inconsistencias", "Inconsistências nos dados"), ("ajustes_de_fonte", "Fonte ajustada"),
                       ("ajustes_de_largura", "Caixa alargada"), ("ajustes_de_altura", "Caixa ampliada")):
        items = sorted(report[key]) if key == "missing_tokens" else report.get(key)
        if items:
            print(f"\n⚠️  {title} ({len(items)}):")
            for m in items:
                print("   -", m)
    print("\n📷", report["photo"])
    if report["missing_tokens"] or report["errors"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
