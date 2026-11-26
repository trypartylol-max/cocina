# Historial de cambios

## 2026-11-25 — Repo propio
- Cocina sale del repo de SHAAH a `trypartylol-max/cocina` (URL nueva: `/cocina/`). La ruta vieja `/shaah/cocina/` redirige aquí.
- Configuración separada en `config.js`; backend versionado en `backend/`.

## 2026-11-24/25 — Historia previa (en el repo de SHAAH, PRs #141–#152)
- Front-end en GitHub Pages con Apps Script como API (login de Google, solo el dueño).
- Lo consumido, perdido o cocinado vuelve a Compras; Compras flash; ✕ Retirar.
- Aviso diario como campana con hora configurable.
- Móvil: pestañas Casa/Compras/Cocinar, deslizar ✓/✕, cifras compactas, Agregar simple, resumen del mes.
- Usar primero (vencido/hoy/mañana/en N días, receta sugerida, botar vencidos, ⏳ Todavía está bueno).
- Inventario con filtros por lugar, cantidad ●●○ y menú por fila; Qué cocinar top 3; KPIs Cocinar ya / Flash.
- Compu: 3 columnas a pantalla completa; tarjetas plegadas con vista previa grande.
- Auto-actualización (version.json) y safe-area en iPhone.

## 2026-11-26 — Mover de lugar y de categoría
- En la ficha de un producto (tocar la fila o ⏳): botones de **Lugar** (Congelador / Nevera / Despensa / Encimera) y selector de **Categoría**.
- El cambio se guarda en el inventario y en el catálogo, así que la próxima compra de ese producto ya entra en su sitio y en compras sale en el pasillo correcto.
- Backend: nueva función `mover(id, ubi, cat)` (versión 25 de la implementación). Copia de `backend/Code.gs` sincronizada con el código vivo (incluye `autorizarDrive` y OCR con 'iw').
