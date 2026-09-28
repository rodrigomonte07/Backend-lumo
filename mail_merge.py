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


def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def apply_field_map(prs, field_map, data, report):
    """Percorre cada regra do field_map e substitui o texto no shape exato."""
    slides = list(prs.slides)

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
                if mode == "exact" and run.text == find:
                    run.text = value
                    applied = True
                elif mode == "contains" and find in run.text:
                    run.text = run.text.replace(find, value)
                    applied = True

        if applied:
            report["applied"] += 1
        else:
            report["not_found"].append(
                f"token {token}: texto '{find}' (mode={mode}) não encontrado no "
                f"shape_id {shape_id} do slide {slide_no} — o texto do template pode "
                f"ter mudado. Verifique field_map.json."
            )


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

    if report["missing_tokens"] or report["errors"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
