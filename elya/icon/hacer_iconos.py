"""Iconos de la app Esquema ELYA: rayo blanco sobre el naranja de la app, con margen para que Android no lo recorte
al ponerlo en círculo (zona segura de los iconos «maskable»).
Uso: python elya/icon/hacer_iconos.py  -> icon-512.png, icon-192.png, icon-180.png, icon-32.png
"""
from pathlib import Path
from PIL import Image, ImageDraw

NARANJA, BLANCO = (196, 86, 29), (255, 255, 255)
RAYO = [(0.56, 0.18), (0.30, 0.55), (0.48, 0.55), (0.42, 0.82), (0.70, 0.43), (0.52, 0.43), (0.60, 0.18)]

big = 1024
im = Image.new("RGB", (big, big), NARANJA)
d = ImageDraw.Draw(im)
d.polygon([(x * big, y * big) for x, y in RAYO], fill=BLANCO)
out = Path(__file__).parent
for n in (512, 192, 180, 32):
    im.resize((n, n), Image.LANCZOS).save(out / f"icon-{n}.png")
