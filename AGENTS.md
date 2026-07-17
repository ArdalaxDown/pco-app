# AGENTS.md — PCO App (Flask + PostgreSQL · Metro L2)

## ⚠️ IMPORTANTE: Responder SIEMPRE en español

## Stack

- **Flask 3.1** (monolithic `app.py`, ~2270 líneas)
- **PostgreSQL** via psycopg2; schema `seguimiento_vias` debles managed outside repo
- **openpyxl** for Excel import/export
- Deployed to **Render.com** via `render.yaml` + `Procfile` (gunicorn)
- DB provisioned at **Neon** (PostgreSQL cloud); URL in `DATABASE_URL` env var

## Dev Workflow

### Two-phase testing (no test runner)

1. **Syntax check**: `python -c "import ast; ast.parse(open('app.py').read())"`
2. **Import check**: `python -c "import app; print('Routes:', len(app.app.url_map.iter_rules()))"`
3. **Functional tests**: write ad-hoc Python scripts with `app.app.test_client()` + mock `get_db_connection` returning `MockConn` / `MockCur`. Template rendering at `/` MUST return 200 before commit.

### Template balance check

```sh
# count opens/closes in the <script> block of index.html (must match)
python -c "
import re; h = open('templates/index.html').read()
s = re.search(r'<script>(.*?)</script>', h, re.DOTALL).group(1)
print('{}=', s.count('{'), '{', s.count('}'))
"
```

Run this after ANY JavaScript edit in `index.html`. A mismatch breaks the entire page silently.

### Auto-git push

After EVERY successful change (code, AGENTS.md, etc.), stage, commit, and push to `origin/main`. Use a concise Spanish commit message describing what was done.

## Auto-logging rule

At the end of each session, update the **Historial de cambios recientes** section below with a bullet-point summary of everything done in this session. This ensures the next AI session has full context without re-reading the conversation.

## Architecture (key sections in app.py)

| Lines | Section                             | FS Notes                                                           |
|-------|-------------------------------------|--------------------------------------------------------------------|
| 22-60 | `CV_MAP` + `ZONE_POSITIONS`         | The spatial registry of the system. Rects in px over `mapa_real.png` (1700×820). |
| 61-200| `ZONE_POSITIONS` (cont)             | ~160 zones; keys like `'E20'`, `'E20->E21 VIA1'`, `'D1'`, `'TK1'` |
| 231-255 | `get_db_connection` (Neon/local)    | Retry logic; calls `SET TIME ZONE 'America/Lima'` per session. |
| 273-290 | `normalizar_zonas` / Safety         | Safety is evaluated intersection-based on `ZONE_POSITIONS` keys. |
| 319-352 | `/` route (index)                   | Filters `WHERE archivado=FALSE AND estado IN ('En Vía','Liberado')` (all days). |
| 431-455 | `_clasificar_zona` (selector cats)  | Must be updated when new zones are added. |
| 464-498 | `/api/zonas_catalogo`               | JSON catalog groups `ZONE_POSITIONS` into categories for frontend selector. |
| 557-600 | `/api/validar_zonas`                | Validates custom zone strings against `ZONE_POSITIONS` + `SINONIMOS`. |
| 1240-1703 | `/importar_excel` + `/confirmar_importado` + `/ingresar` + `/importar_texto` | OT duplication check, VARCHAR truncation, text import. |
| 2272 | EOF                                 |                                                                   |

## DB Constraints

- Tables were created as `VARCHAR(150)` for most fields (no SQL migration in repo).
- `LIMITE_CAMPO = 150` and `_truncar()` protect all insert paths. The truncation happens in-memory before `cur.execute(INSERT)`.
- Recommended SQL migration (NOT applied in repo): convert `VARCHAR` to `TEXT`:
  ```sql
  ALTER TABLE seguimiento_vias
    ALTER COLUMN responsable      TYPE TEXT,
    ALTER COLUMN comentario       TYPE TEXT,
    ALTER COLUMN orden_trabajo    TYPE TEXT,
    ALTER COLUMN ubicacion_zona   TYPE TEXT;
  ```

## Localization quirk

- `extract_tetra()` normalizes text separators (commas, spaces, colons) looking for `22\d{3}` tokens.
- `parse_hora_flexible()` accepts `HH`, `HH:MM`, and `HH:MM:SS` (no-required format).
- Timezone: always `America/Lima` (`SET TIME ZONE` on each DB connection).

## The Zone Selector (🚧 complex)

Path: **`/api/zonas_catalogo` → `index.html` modal `#modalSelectorZona`**

- Category mapping lives in `_clasificar_zona` (`app.py:431`). Changing keys or adding new zones here affects:
  - `api_zonas_catalogo` response structure → `CATEGORIAS_VISIBLES` (app.py:486)
  - Frontend boton filter buttons: `data-filtro-mapa="..."` in `#filtroMapaGrupo` (start line ~188)
  - Chips rendering: `titulos[cat]` inside `renderChipsPorCategoria()` (line ~1777)
  - Popover category label inside `abrirPopoverZona()` — `catLabel` object (line ~1717)
- When **editing a zone coordinate** (rect), look at `ZONE_POSITIONS` dict (`app.py:61+`).
  - Use `/coordenadas` page (intact) to preview overlay rects and find `{top,left,width,height}`.
  - Then edit `ZONE_POSITIONS` dict (the key name is what the user writes in `ubicacion_zona`; the rect data is what `mapa.html` renders as overlays).
- When a user writes custom aliases (`ESTACION22 → E22`), add synonym to `SINONIMOS` dict in `api_validar_zonas` (app.py:553).

### Mini-map modes

- **Default (`todos`/`estacion`/`pozo`/`ptsa`/`deposito`)** → `renderMiniMapa` draws rectangle areas (`.zona-select-area`). Click triggers `onClicAreaZona` → popover if overlap.
- **Tramo filter (`tramo`/`tramo_via`)** → switches to `renderMiniMapaNodos` (node circles `E20-E24`, `PV19-PV24`). Click = `onClicNodoTramo` with `selectorZonaState.tramoOrigenSel`.
- `ptsa` category includes `TK1-TK7`, `D1`, `D2`, `PTSA`.

### Modal stacking (important)

- When `#modalSelectorZona` opens from a `editModal` (Corregir), JS explicitly sets `z-index` of the parent `editModal` to `1045` and the selector to `1070` plus `backdrop:static`. Without this, clicking in the mini-map steal the focus, Corregir closes silently, and the user cannot apply zone correction.
- The `abrirSelectorZona` function (`index.html` JS) detects if the input belongs to a `editModal` and applies these z-index + backdrop static.

## OT Duplicate validation

- `/ingresar` and `/confirmar_importado` each execute a `SELECT ... WHERE orden_trabajo ILIKE ... AND hora_fin IS NULL` before insertion.
- AJAX endpoint `/verificar_ot_duplicada` provides real-time validation in the registration form.
- When `confirmar_importado` returns `duplicada:true`, the frontend JSON handler restores the checkbox and keeps the row in the preview table (does NOT delete the pending row from session).

## UI-supported zone aliases (SINONIMOS)

- Defined once in `api_validar_zonas` (app.py:553+). Example: `ESTACION22 -> E22`, `CAJA TIPO 1 RAMAL D1 -> D1`.
- New alias: add to `SINONIMOS` dict and verify with `POST /api/validar_zonas` (test client).
- Currently (~17 entries). If a new Excel format is detected, add its mapping here.

## Historial de cambios recientes

- **Monitoreo persistente**: Cambio en la query de `/` (index) para mostrar TODOS los trabajos no archivados con estado 'En Vía' o 'Liberado', sin filtro de fecha. Así los trabajos finalizados permanecen visibles en la ventana de monitoreo hasta que el usuario pulsa "Archivar Turno".
- **NTA/NTP**: Agregadas a `ZONE_POSITIONS` (app.py:201-202) y clasificadas como `ptsa` en `_clasificar_zona`.
- **Filtro chips/popover por categoría**: `renderChipsPorCategoria` y `onClicAreaZona` ahora respetan `filtroMapa` — en modo "Estaciones" solo se ven/eligen estaciones, sin tramos mezclados.
- **Acceso a vía en importación**: Tanto en `importar_excel` como `importar_texto` y `_parsear_fila_fija`, ahora se salta filas solo si `acceso == 'NO'`; antes solo aceptaba `'SI'`. `estado` debe ser `'AUTORIZADA'` (antes era `'CONFIRMADA'`).
- **Respaldo DB**: Se sugirió `pg_dump` diario como backup de Neon. Tener clon local del repo para emergencias si Render cae.
## Migración BD requerida (Neon / PostgreSQL)

Para soportar zonas peatonales + biviales en el mismo permiso, ejecutar:

```sql
ALTER TABLE seguimiento_vias 
ADD COLUMN IF NOT EXISTS ubicacion_zona_peatonal TEXT;
```
