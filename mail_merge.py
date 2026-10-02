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
import shutil
import sys
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu, Pt

EMU_PER_PT = 12700

# Fração do tamanho da fonte (em pt) usada como largura média de caractere.
# É uma aproximação (não medimos glifo por glifo) calibrada para as fontes
# bold/condensadas deste template — suficiente para decidir COM MARGEM DE
# SOBRA se um texto cabe, sem precisar renderizar a fonte de verdade.
AVG_CHAR_WIDTH_FACTOR = 0.64
SAFETY_MARGIN = 0.88          # deixa ~6% de folga na largura utilizável
MIN_SHRINK_RATIO = 0.70       # nunca encolhe a fonte abaixo de 70% do original
LINE_HEIGHT_FACTOR = 1.22     # aproximação de entrelinha quando não há lnSpc


def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _line_capacity_chars(box_width_emu, font_size_pt):
    """Quantos caracteres cabem em uma linha, dada a largura da caixa e o tamanho da fonte."""
    usable_width_pt = (box_width_emu / EMU_PER_PT) * SAFETY_MARGIN
    char_width_pt = max(font_size_pt * AVG_CHAR_WIDTH_FACTOR, 0.1)
    return max(1, int(usable_width_pt / char_width_pt))


def _wrapped_line_count(text, capacity_chars):
    """
    Simula quebra de linha por palavra (como o PowerPoint faz) e conta quantas
    linhas resultam. Trata tambem o caso de uma palavra sozinha (sem espaco)
    ser mais larga que a propria caixa - um numero grande tipo
    "1.139.900,75" nao tem espaco nenhum, e quando nao cabe o PowerPoint
    quebra letra por letra (produzindo um digito "orfao" numa linha), em vez
    de simplesmente deixar a palavra inteira estourar. Sem tratar esse caso
    aqui, o calculo achava que a palavra "sempre cabe" e nunca detectava a
    necessidade de encolher a fonte.
    """
    if capacity_chars <= 0:
        return 1
    words = text.split(" ")
    lines, cur = 1, 0
    for w in words:
        wlen = len(w)
        if wlen > capacity_chars:
            if cur > 0:
                lines += 1
                cur = 0
            lines += wlen // capacity_chars
            cur = wlen % capacity_chars
            continue
        add = wlen + (1 if cur > 0 else 0)
        if cur + add > capacity_chars and cur > 0:
            lines += 1
            cur = wlen
        else:
            cur += add
    return max(1, lines)


def fit_run_to_shape(shape, run, original_text, slide=None, report=None):
    """
    Garante que o texto final do `run` caiba na caixa sem estourar o espaco
    que o desenho original reservava para ele.

    Por que: toda caixa deste template usa `spAutoFit` (a altura deveria se
    ajustar ao texto). O PowerPoint e o Canva recalculam isso ao vivo, mas
    nem todo renderizador faz isso (o LibreOffice, por exemplo, usa a altura
    congelada no arquivo) - entao um valor mais longo que o texto de exemplo
    original pode quebrar linha e invadir o elemento vizinho, ou ser cortado.

    Estrategia: 1) usa o PROPRIO texto original (o que estava no arquivo
    antes da troca) como referencia de quantas linhas a caixa foi desenhada
    para comportar - mais confiavel do que calcular so a partir da altura em
    EMU, porque cancela qualquer erro de calibracao da formula de largura de
    caractere. 2) tenta encolher a fonte (ate um piso) para caber dentro
    dessa referencia. 3) Se mesmo no piso nao couber, cresce a altura da
    caixa E empurra para baixo, na mesma coluna, qualquer elemento que
    estivesse logo abaixo - para nao invadir o proximo card/titulo.
    """
    if run.font.size is None:
        return
    original_sz_pt = run.font.size.pt
    box_width_emu = shape.width
    box_height_emu = shape.height
    if not box_width_emu or not box_height_emu:
        return

    text = run.text
    capacity_chars_original = _line_capacity_chars(box_width_emu, original_sz_pt)
    baseline_lines = max(1, _wrapped_line_count(original_text, capacity_chars_original))

    sz = original_sz_pt
    for _ in range(14):
        capacity_chars = _line_capacity_chars(box_width_emu, sz)
        if _wrapped_line_count(text, capacity_chars) <= baseline_lines:
            break
        sz *= 0.96
        if sz <= original_sz_pt * MIN_SHRINK_RATIO:
            sz = original_sz_pt * MIN_SHRINK_RATIO
            break

    # Segunda camada de seguranca, independente da calibracao de largura por
    # caractere (que nunca vai ser 100% precisa sem medir a fonte de
    # verdade): se o maior "token" sem espaco do texto novo (ex: um valor
    # monetario) for mais comprido que o do texto original, garante uma
    # reducao PROPORCIONAL minima de fonte baseada so na contagem de
    # caracteres. Isso pega casos-limite que o calculo de capacidade por si
    # so as vezes erra por 1 caractere (arredondamento).
    longest_original_token = max((len(w) for w in original_text.split(" ")), default=1)
    longest_new_token = max((len(w) for w in text.split(" ")), default=1)
    if longest_new_token > longest_original_token > 0:
        ratio_sz = original_sz_pt * (longest_original_token / longest_new_token)
        sz = min(sz, max(ratio_sz, original_sz_pt * MIN_SHRINK_RATIO))

    if sz < original_sz_pt - 0.05:
        run.font.size = Pt(round(sz, 1))
        if report is not None:
            report.setdefault("ajustes_de_fonte", []).append(
                f"shape_id {shape.shape_id}: fonte reduzida de {original_sz_pt:.1f}pt "
                f"para {sz:.1f}pt para caber '{text[:40]}{'...' if len(text) > 40 else ''}'"
            )

    capacity_chars = _line_capacity_chars(box_width_emu, sz)
    lines_needed = _wrapped_line_count(text, capacity_chars)
    if lines_needed > baseline_lines:
        extra_lines = lines_needed - baseline_lines
        line_height_pt = sz * LINE_HEIGHT_FACTOR
        extra_height_emu = int(extra_lines * line_height_pt * EMU_PER_PT)
        original_bottom = shape.top + box_height_emu
        shape.height = Emu(box_height_emu + extra_height_emu)
        if report is not None:
            report.setdefault("ajustes_de_altura", []).append(
                f"shape_id {shape.shape_id}: caixa aumentada em {extra_lines} linha(s) "
                f"para nao cortar '{text[:40]}{'...' if len(text) > 40 else ''}'"
            )
        if slide is not None:
            _cascade_shift_below(slide, shape, original_bottom, extra_height_emu, report)


def _cascade_shift_below(slide, resized_shape, original_bottom_emu, delta_emu, report=None):
    """
    Empurra para baixo, na mesma coluna, qualquer shape que estava posicionado
    logo abaixo da caixa que acabou de crescer - para a caixa maior nao passar
    por cima do que vinha depois dela (ex: o titulo crescer e cobrir o
    "Proximos passos" e os cards de baixo).

    "Mesma coluna" = sobreposicao horizontal de pelo menos 30% com a caixa
    redimensionada. Isso evita empurrar elementos que so por coincidencia tem
    o topo mais baixo mas ficam ao lado (ex: a foto da escola a direita).
    """
    if delta_emu <= 0:
        return
    rx1 = resized_shape.left
    rx2 = resized_shape.left + resized_shape.width
    tolerance = 9525 * 4  # ~4px de tolerancia para posicoes "coladas"

    for other in slide.shapes:
        if other.shape_id == resized_shape.shape_id:
            continue
        if other.left is None or other.width is None or other.top is None:
            continue
        ox1 = other.left
        ox2 = other.left + other.width
        overlap = min(rx2, ox2) - max(rx1, ox1)
        min_width = min(resized_shape.width, other.width)
        if min_width <= 0 or overlap / min_width < 0.3:
            continue
        if other.top + tolerance >= original_bottom_emu:
            other.top = Emu(int(other.top) + delta_emu)
            if report is not None:
                report.setdefault("elementos_reposicionados", []).append(
                    f"shape_id {other.shape_id} deslocado {delta_emu / EMU_PER_PT / 72:.2f}in "
                    f"para baixo (por causa do crescimento do shape_id {resized_shape.shape_id})"
                )


def apply_field_map(prs, field_map, data, report):
    """Percorre cada regra do field_map e substitui o texto no shape exato."""
    slides = list(prs.slides)
    _original_run_text = {}   # id(run) -> texto original antes de qualquer troca
    _touched_runs = {}        # id(run) -> (shape, run, slide) tocados nesta chamada

    for rule in field_map:
        token = rule["token"]
        slide_no = rule["slide"]  # 1-based, igual ao nome do slideN.xml
        shape_id = rule["shape_id"]
        mode = rule["mode"]
        find = rule["find"]

        if token not in data:
            report["missing_tokens"].add(token)
            continue

        value = str(data[token])

        if slide_no < 1 or slide_no > len(slides):
            report["errors"].append(f"slide {slide_no} não existe (regra do token {token})")
            continue

        slide = slides[slide_no - 1]
        shape = _find_shape_by_id(slide.shapes, shape_id)
        if shape is None:
            report["errors"].append(
                f"shape_id {shape_id} não encontrado no slide {slide_no} (token {token})"
            )
            continue
        if not shape.has_text_frame:
            report["errors"].append(
                f"shape_id {shape_id} no slide {slide_no} não é uma caixa de texto (token {token})"
            )
            continue

        applied = False
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                # Guarda o texto ORIGINAL do run na primeira vez que o tocamos
                # nesta chamada (mesmo shape pode ser alvo de varias regras do
                # field_map - ex: nome da escola E percentual na mesma frase -
                # e precisamos de UM UNICO ajuste de tamanho no final, nao um
                # por regra, senao a fonte encolhe em cascata sem necessidade).
                # id(run) NAO e estavel entre iteracoes: o python-pptx cria um
                # objeto wrapper novo a cada acesso a .paragraphs/.runs, mesmo
                # apontando para o MESMO elemento XML. Usamos id(run._r) - o
                # elemento lxml real por baixo - que sim e estavel, senao duas
                # regras que tocam o mesmo texto (ex: nome + percentual na
                # mesma frase) disparam o ajuste de fonte duas vezes em cascata.
                run_key = id(run._r)
                if run_key not in _original_run_text:
                    _original_run_text[run_key] = run.text

                changed = False
                if mode == "exact" and run.text == find:
                    run.text = value
                    changed = True
                elif mode == "contains" and find in run.text:
                    run.text = run.text.replace(find, value)
                    changed = True
                if changed:
                    applied = True
                    _touched_runs[run_key] = (shape, run, slide)

        if applied:
            report["applied"] += 1
        else:
            report["not_found"].append(
                f"token {token}: texto '{find}' (mode={mode}) não encontrado no "
                f"shape_id {shape_id} do slide {slide_no} — o texto do template pode "
                f"ter mudado. Verifique field_map.json."
            )

    # Uma unica passada de ajuste de tamanho por run, depois que TODAS as
    # substituicoes desta chamada ja foram aplicadas - evita encolher a
    # mesma caixa varias vezes quando duas regras tocam o mesmo texto.
    for run_key, (shape, run, slide) in _touched_runs.items():
        fit_run_to_shape(shape, run, _original_run_text[run_key], slide=slide, report=report)


def _find_shape_by_id(shapes, target_id):
    """Busca recursiva (inclusive dentro de grupos) por shape_id."""
    for shape in shapes:
        if shape.shape_id == target_id:
            return shape
        if shape.shape_type == 6:  # GROUP
            found = _find_shape_by_id(shape.shapes, target_id)
            if found is not None:
                return found
    return None


def replace_photo(pptx_path_out, photo_path, report, media_filename="image14.png"):
    """
    Troca a foto da fachada da escola.

    No template de referência, a mesma foto (image14.png) é reaproveitada
    em 3 lugares (slides 2, 3 e 9) via o MESMO arquivo de mídia dentro do
    .pptx — então basta substituir esse único arquivo dentro do .zip que
    as 3 aparições são atualizadas de uma vez, sem tocar em nenhum shape.
    """
    if not photo_path:
        report["photo"] = "nenhuma foto fornecida — mantida a foto original do template"
        return

    # O slot de mídia é .png — reconvertemos qualquer formato de entrada (jpg, webp...)
    # para PNG de verdade, para não deixar a extensão e os bytes reais dessincronizados
    # (o que o PowerPoint costuma marcar como arquivo corrompido / pedir reparo).
    from io import BytesIO
    from PIL import Image

    with Image.open(photo_path) as im:
        im = im.convert("RGB")
        buf = BytesIO()
        im.save(buf, format="PNG")
        photo_bytes = buf.getvalue()

    # Reabre o pptx (que já foi salvo com os textos) como zip e sobrescreve a mídia
    tmp_path = str(pptx_path_out) + ".tmp"
    with zipfile.ZipFile(pptx_path_out, "r") as zin:
        names = zin.namelist()
        target = f"ppt/media/{media_filename}"
        if target not in names:
            report["errors"].append(
                f"Não encontrei ppt/media/{media_filename} no arquivo final — "
                f"a foto NÃO foi trocada. Confira o nome do arquivo de mídia no "
                f"template (pode mudar se o Canva reexportar o arquivo)."
            )
            return
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = photo_bytes if item.filename == target else zin.read(item.filename)
                zout.writestr(item, data)

    shutil.move(tmp_path, pptx_path_out)
    report["photo"] = f"foto substituída em ppt/media/{media_filename} (usada nos slides 2, 3 e 9)"


def main():
    ap = argparse.ArgumentParser(description="Mail-merge para o template Rede Lumo (Canva)")
    ap.add_argument("--template", required=True, help="Caminho do .pptx original (Canva)")
    ap.add_argument("--data", required=True, help="JSON com token -> valor")
    ap.add_argument("--map", default="field_map.json", help="JSON com as regras de substituição")
    ap.add_argument("--foto", default=None, help="Caminho da foto da fachada da escola (opcional)")
    ap.add_argument("--out", required=True, help="Caminho do .pptx de saída")
    args = ap.parse_args()

    field_map = load_json(args.map)
    data = load_json(args.data)

    prs = Presentation(args.template)

    report = {
        "applied": 0,
        "not_found": [],
        "missing_tokens": set(),
        "errors": [],
        "photo": None,
    }

    apply_field_map(prs, field_map, data, report)
    prs.save(args.out)
    replace_photo(args.out, args.foto, report)

    # Relatório de validação — é assim que você audita se ficou 100% fiel
    print(f"\n✅ {report['applied']} campos substituídos com sucesso.")
    if report["missing_tokens"]:
        print(f"\n⚠️  Tokens usados no field_map mas ausentes em data.json "
              f"({len(report['missing_tokens'])}):")
        for t in sorted(report["missing_tokens"]):
            print(f"   - {t}")
    if report["not_found"]:
        print(f"\n⚠️  Regras que não encontraram o texto esperado "
              f"({len(report['not_found'])}):")
        for msg in report["not_found"]:
            print(f"   - {msg}")
    if report["errors"]:
        print(f"\n❌ Erros ({len(report['errors'])}):")
        for msg in report["errors"]:
            print(f"   - {msg}")
    if report["photo"]:
        print(f"\n📷 {report['photo']}")
    if report.get("ajustes_de_fonte"):
        print(f"\n🔧 Fonte ajustada automaticamente para caber ({len(report['ajustes_de_fonte'])}):")
        for msg in report["ajustes_de_fonte"]:
            print(f"   - {msg}")
    if report.get("ajustes_de_altura"):
        print(f"\n🔧 Caixa ajustada automaticamente para caber ({len(report['ajustes_de_altura'])}):")
        for msg in report["ajustes_de_altura"]:
            print(f"   - {msg}")

    if report["missing_tokens"] or report["errors"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
