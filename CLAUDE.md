# CLAUDE.md — Cocina

App personal de inventario de cocina. Proyecto **independiente** de SHAAH: no usa la Bitácora de SHAAH ni sus tests. El historial va en `CHANGELOG.md`.

## Estructura
- `index.html` (HTML + CSS), `app.js` (lógica), `api.js` (puente con la API: imita `google.script.run` con `GS()`), `config.js` (URL de la API de esta instalación).
- `vendor/supabase-2.118.0.js` (login; mismo archivo y SRI que SHAAH).
- `backend/Code.gs` + `backend/appsscript.json`: copia del proyecto Apps Script **"Cocina web"** (id `1z9tZ1Y3RwaK2qMknvn2SicfGnzugtt1PVC3FD6sGvf3KhGCdNgEU-Jdm`), sobre la hoja "Cocina · Inventario v1". El código vivo está en Apps Script; mantén esta copia igual tras cada cambio.

## Backend (Apps Script)
- `doPost` recibe `{jwt, fn, args}`, valida el JWT de Supabase contra `/auth/v1/user` y exige `OWNER`. Solo llama funciones de la lista `API`.
- Implementación web: ejecutar como el dueño, acceso "Cualquiera". Para publicar un cambio: editar en el editor de Apps Script → guardar → Implementar → Gestionar implementaciones → editar → **Nueva versión** → Implementar (la URL no cambia).
- Columnas añadidas sobre la marcha con `ensureCol_`: `Flash` (Lista_de_Compras), `Retirado` (Catalogo), `Vence_nuevo` / `Usar_antes` (Inventario).

## Publicar el front-end
- GitHub Pages sirve `main`. El propietario revisa antes: no hagas merge de un cambio visible sin su "publica".
- **Cache-busting obligatorio** en cada cambio: actualiza `?v=<8 primeros del sha256>` de `app.js`/`api.js` en `index.html`, y pon la misma marca nueva en `<meta name="cocina-v">` y en `version.json` (la app compara ambas y se recarga sola, también en el icono del iPhone).
- Verifica el despliegue comparando hashes de lo publicado.

## Diseño
- Móvil (≤900 px): pestañas abajo (Casa / Compras / Cocinar), cifras en una fila, deslizar ✓/✕.
- Compu (>900 px): 3 columnas a pantalla completa con scroll propio; tarjetas plegadas con vista previa grande. Se juzga a 2560 px.
