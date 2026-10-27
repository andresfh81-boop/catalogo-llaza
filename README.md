# Catálogo Llaza (app instalable)

App instalable (PWA) del Catálogo Llaza: el mismo aspecto que la app de Google Apps Script (repositorio privado `catalogo-llaza-app`), pero instalable de verdad, a pantalla completa y con la última copia guardada para consultar sin conexión.

Enlace: https://andresfh81-boop.github.io/catalogo-llaza/

- `index.html`, `app/app.css` y `app/client.js` se generan con `python build_pwa.py` en `catalogo-llaza-app`: no editarlos aquí.
- `app/pwa.js`: inicio de sesión con Google, llamadas a la API (proyecto «Catálogo Llaza API») y copia sin conexión. `app/pwa.css`: lo que se oculta en esta fase (editar, añadir, hojas, exportar).
- `sw.js`: guarda la app para abrirla sin conexión. Al cambiar la lista de archivos, subir `V`.
- Icono: el logo de Google Apps Script (el mismo del acceso directo del escritorio), generado con `python icon/hacer_iconos.py`.

Fase 1: consultar (buscar, fichas, fotos, documentación, buscar con foto). Para editar: botón «✏ App completa». Fases siguientes en el README de `catalogo-llaza-app`.

## App aparte: Esquema eléctrico ELYA

Enlace: https://andresfh81-boop.github.io/catalogo-llaza/elya/ (se instala como app independiente, con su icono, y funciona sin conexión).

- Carpeta `elya/`: `index.html` se genera con `python build_elya.py` en `catalogo-llaza-app` (a partir de `esquema/elya.html`): no editarlo aquí.
- `elya/sw.js` usa cachés «elya-…» y el `sw.js` del catálogo solo borra las «cl-…», para que una app no borre la copia sin conexión de la otra.
- `elya/lib/pdfjs/`: pdf.js 3.11.174 (npm pdfjs-dist, licencia Apache 2.0) para leer el PDF del plano sin conexión.
- Icono: rayo blanco sobre naranja, generado con `python elya/icon/hacer_iconos.py`.

## Interruptor remoto (`bloqueo.json`)

Lo leen Modulbox-400 (`modulbox`), Sombrex (`sombrex`), Kuadbox (`kuadbox`) y las dos apps de etiquetas (`etiquetas`). Con `"bloqueado": true` (o el modelo en `"modelos"`: OPEN, SemiOPEN, SemiBOX, FullBOX) la app deja de funcionar sin ningún aviso: en los configuradores el plano no se ve e Imprimir/PDF/PNG no hacen nada; en etiquetas la hoja de corte no se lee. Sin conexión vale el último estado leído. Se edita desde github.com o la app de GitHub y tarda 1-2 minutos.

`"llave"`: SHA-256 del PBKDF2-SHA256 (sal `llaza-bloqueo`, 200 000 vueltas) del código de Andrés. El aparato guarda el PBKDF2 y lo publicado no sirve para fabricarlo. Se mete el código con 5 toques seguidos en el título del configurador o en «Subir hoja de corte»; si no coincide, la app enseña la llave para copiarla aquí. El aparato con la llave buena no se bloquea nunca (en los configuradores sale 🔓 abajo a la derecha). En las apps web la llave vale para todas a la vez en ese navegador (mismo dominio); en etiquetas de escritorio se guarda en `.llave`. Cambiar la llave invalida todas las anteriores.
