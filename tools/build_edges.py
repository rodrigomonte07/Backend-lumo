#!/usr/bin/env python3
"""
Gera as faixas decorativas das laterais (esquerda/direita) de cada slide como
UMA imagem só por lado, a partir do PDF exportado do Canva (referência vetorial).

Por quê: as laterais do template são ~10 formas livres com rotação negativa +
flip, que o LibreOffice desenha errado (daí as "peças soltas" no PDF). Uma imagem
por lado fica idêntica ao Canva em qualquer renderizador (PowerPoint, Keynote,
Google Slides, LibreOffice).

Uso:  python3 tools/build_edges.py assets/modelo_canva_referencia.pdf assets/edges
Requer: poppler-utils (pdftoppm) e Pillow.
"""
import subprocess, sys, tempfile
from pathlib import Path
from PIL import Image

SLIDE_W_EMU = 18288000
DPI = 300                      # 1440pt * 300/72 = 6000 px de largura
LEFT_EDGE_EMU = 800000         # tudo que termina antes disso é "lateral esquerda"
RIGHT_EDGE_EMU = 16900000      # tudo que começa depois disso é "lateral direita"
PAGES = range(2, 10)           # slides 2..9 (capa e fecho têm decoração própria)

def main(pdf, outdir):
    outdir = Path(outdir); outdir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["pdftoppm", "-r", str(DPI), "-png", pdf, f"{tmp}/p"], check=True)
        for n in PAGES:
            im = Image.open(f"{tmp}/p-{n:02d}.png").convert("RGB")
            W, H = im.size
            emu_px = SLIDE_W_EMU / W
            lw = int(-(-LEFT_EDGE_EMU // emu_px))          # ceil
            rx = int(RIGHT_EDGE_EMU // emu_px)             # floor
            im.crop((0, 0, lw, H)).save(outdir / f"s{n}_L.png", optimize=True)
            im.crop((rx, 0, W, H)).save(outdir / f"s{n}_R.png", optimize=True)
            print(n, "L", lw, "px  R", W - rx, "px  EMU/px", round(emu_px, 2))

def corners(pdf, outdir):
    """Capa (1) e fecho (10): a peça laranja do canto direito é outra forma que o LibreOffice desenha torta."""
    X0, Y0, Y1 = 17732542, 2205962, 4359563
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["pdftoppm", "-r", str(DPI), "-png", "-f", "1", "-l", "10", pdf, f"{tmp}/p"], check=True)
        for n in (1, 10):
            im = Image.open(f"{tmp}/p-{n:02d}.png").convert("RGB")
            e = SLIDE_W_EMU / im.size[0]
            im.crop((int(X0 // e), int(Y0 // e), im.size[0], int(Y1 // e))).save(Path(outdir) / f"s{n}_corner.png", optimize=True)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
    corners(sys.argv[1], sys.argv[2])
