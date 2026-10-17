"""Iconos de la app: el mismo logo de Google Apps Script que sale en el acceso directo del escritorio.
Fondo blanco y margen para que Android no recorte el logo al ponerlo en círculo.
Uso: python icon/hacer_iconos.py  -> icon/icon-512.png, icon-192.png, icon-180.png, icon-32.png
"""
import io
import urllib.request
from pathlib import Path
from PIL import Image

LOGO = "https://www.gstatic.com/images/branding/product/1x/apps_script_512dp.png"

logo = Image.open(io.BytesIO(urllib.request.urlopen(LOGO).read())).convert("RGBA")
out = Path(__file__).parent
for n in (512, 192, 180):
    fondo = Image.new("RGBA", (n, n), "white")
    lado = round(n * 0.72)  # dentro de la zona segura de los iconos «maskable»
    fondo.alpha_composite(logo.resize((lado, lado), Image.LANCZOS), ((n - lado) // 2, (n - lado) // 2))
    fondo.convert("RGB").save(out / f"icon-{n}.png")
logo.resize((32, 32), Image.LANCZOS).save(out / "icon-32.png")
