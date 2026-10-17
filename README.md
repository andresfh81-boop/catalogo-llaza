# Catálogo Llaza (app instalable)

App instalable (PWA) del Catálogo Llaza: el mismo aspecto que la app de Google Apps Script (repositorio privado `catalogo-llaza-app`), pero instalable de verdad, a pantalla completa y con la última copia guardada para consultar sin conexión.

Enlace: https://andresfh81-boop.github.io/catalogo-llaza/

- `index.html`, `app/app.css` y `app/client.js` se generan con `python build_pwa.py` en `catalogo-llaza-app`: no editarlos aquí.
- `app/pwa.js`: inicio de sesión con Google, llamadas a la API (proyecto «Catálogo Llaza API») y copia sin conexión. `app/pwa.css`: lo que se oculta en esta fase (editar, añadir, hojas, exportar).
- `sw.js`: guarda la app para abrirla sin conexión. Al cambiar la lista de archivos, subir `V`.
- Icono: el logo de Google Apps Script (el mismo del acceso directo del escritorio), generado con `python icon/hacer_iconos.py`.

Fase 1: consultar (buscar, fichas, fotos, documentación, buscar con foto). Para editar: botón «✏ App completa». Fases siguientes en el README de `catalogo-llaza-app`.
