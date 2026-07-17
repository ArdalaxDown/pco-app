# AGENTS.md — PCO App (Flask + PostgreSQL · Metro L2)

## ⚠️ IMPORTANTE: Responder SIEMPRE en español

## Stack

- **Flask 3.1** (monolítico `app.py`, ~2270 líneas)
- **PostgreSQL** vía psycopg2; esquema `seguimiento_vias` gestionado fuera del repo
- **openpyxl** para importación/exportación Excel
- Desplegado en **Render.com** vía `render.yaml` + `Procfile` (gunicorn)
- BD provisionada en **Neon** (PostgreSQL cloud); URL en variable de entorno `DATABASE_URL`

## Flujo de Desarrollo

### Pruebas en dos fases (sin test runner)

1. **Chequeo de sintaxis**: `python -c "import ast; ast.parse(open('app.py').read())"`
2. **Chequeo de importación**: `python -c "import app; print('Rutas:', len(list(app.app.url_map.iter_rules())))"`
3. **Pruebas funcionales**: escribir scripts Python ad-hoc con `app.app.test_client()` + mock de `get_db_connection` devolviendo `MockConn` / `MockCur`. El renderizado de la plantilla en `/` DEBE retornar 200 antes de commit.

### Chequeo de balance de plantillas

```sh
# contar aperturas/cierres en el bloque <script> de index.html (deben coincidir)
python -c "
import re; h = open('templates/index.html').read()
s = re.search(r'<script>(.*?)</script>', h, re.DOTALL).group(1)
print('{}:', s.count('{'), '}:', s.count('}'))
"
```

Ejecutar esto tras CUALQUIER edición de JavaScript en `index.html`. Un desbalance rompe la página entera silenciosamente.

### Auto-git push

Tras CADA cambio exitoso (código, AGENTS.md, etc.), hacer stage, commit y push a `origin/main`. Usar mensaje de commit conciso en español describiendo lo hecho.

## Regla de auto-logging

Al final de cada sesión, actualizar la sección **Historial de cambios recientes** abajo con un resumen en viñetas de todo lo hecho en la sesión. Esto asegura que la próxima sesión de IA tenga contexto completo sin releer la conversación.

## Arquitectura (secciones clave en app.py)

| Líneas | Sección                              | Notas FS                                                           |
|--------|--------------------------------------|--------------------------------------------------------------------|
| 22-60  | `CV_MAP` + `ZONE_POSITIONS`          | Registro espacial del sistema. Rects en px sobre `mapa_real.png` (1700×820). |
| 61-200 | `ZONE_POSITIONS` (cont)              | ~160 zonas; claves como `'E20'`, `'E20->E21 VIA1'`, `'D1'`, `'TK1'` |
| 231-255 | `get_db_connection` (Neon/local)     | Reintentos; llama `SET TIME ZONE 'America/Lima'` por sesión. |
| 273-290 | `normalizar_zonas` / Safety          | Safety evaluado por intersección basado en claves de `ZONE_POSITIONS`. |
| 319-352 | `/` route (index)                    | Filtra `WHERE archivado=FALSE AND estado IN ('En Vía','Liberado')` (todos los días). |
| 431-455 | `_clasificar_zona` (cats selector)   | Actualizar cuando se añadan nuevas zonas. |
| 464-498 | `/api/zonas_catalogo`                | Catálogo JSON agrupa `ZONE_POSITIONS` en categorías para selector frontend. |
| 557-600 | `/api/validar_zonas`                 | Valida strings de zona custom contra `ZONE_POSITIONS` + `SINONIMOS`. |
| 1240-1703 | `/importar_excel` + `/confirmar_importado` + `/ingresar` + `/importar_texto` | Check duplicados OT, truncamiento VARCHAR, importación texto. |
| 2272   | EOF                                  |                                                                   |

## Restricciones BD

- Tablas creadas como `VARCHAR(150)` para la mayoría de campos (no hay migración SQL en repo).
- `LIMITE_CAMPO = 150` y `_truncar()` protegen todos los paths de inserción. El truncamiento ocurre en memoria antes del `cur.execute(INSERT)`.
- Migración SQL recomendada (NO aplicada en repo): convertir `VARCHAR` a `TEXT`:
  ```sql
  ALTER TABLE seguimiento_vias
    ALTER COLUMN responsable      TYPE TEXT,
    ALTER COLUMN comentario       TYPE TEXT,
    ALTER COLUMN orden_trabajo    TYPE TEXT,
    ALTER COLUMN ubicacion_zona   TYPE TEXT;
  ```

## Particularidades de localización

- `extract_tetra()` normaliza separadores de texto (comas, espacios, dos puntos) buscando tokens `22\d{3}`.
- `parse_hora_flexible()` acepta `HH`, `HH:MM`, y `HH:MM:SS` (formato no obligatorio).
- Zona horaria: siempre `America/Lima` (`SET TIME ZONE` en cada conexión BD).

## El Selector de Zonas (🚧 complejo)

Ruta: **`/api/zonas_catalogo` → `index.html` modal `#modalSelectorZona`**

- Mapeo de categorías vive en `_clasificar_zona` (`app.py:431`). Cambiar claves o añadir zonas aquí afecta:
  - Estructura respuesta `api_zonas_catalogo` → `CATEGORIAS_VISIBLES` (app.py:486)
  - Botones filtro frontend: `data-filtro-mapa="..."` en `#filtroMapaGrupo` (inicio línea ~188)
  - Render chips: `titulos[cat]` dentro de `renderChipsPorCategoria()` (línea ~1777)
  - Etiqueta categoría popover dentro de `abrirPopoverZona()` — objeto `catLabel` (línea ~1717)
- Al **editar coordenada de zona** (rect), ver `ZONE_POSITIONS` dict (`app.py:61+`).
  - Usar página `/coordenadas` (intacta) para previsualizar rects overlay y encontrar `{top,left,width,height}`.
  - Luego editar dict `ZONE_POSITIONS` (el nombre de la clave es lo que el usuario escribe en `ubicacion_zona`; los datos del rect son lo que `mapa.html` renderiza como overlays).
- Cuando usuario escribe alias custom (`ESTACION22 → E22`), añadir sinónimo a dict `SINONIMOS` en `api_validar_zonas` (app.py:553).

### Modos mini-mapa

- **Default (`todos`/`estacion`/`pozo`/`ptsa`/`deposito`)** → `renderMiniMapa` dibuja áreas rectangulares (`.zona-select-area`). Clic dispara `onClicAreaZona` → popover si hay solape.
- **Filtro tramo (`tramo`/`tramo_via`)** → cambia a `renderMiniMapaNodos` (círculos nodos `E20-E24`, `PV19-PV24`). Clic = `onClicNodoTramo` con `selectorZonaState.tramoOrigenSel`.
- Categoría `ptsa` incluye `TK1-TK7`, `D1`, `D2`, `PTSA`.

### Apilado de modales (importante)

- Cuando `#modalSelectorZona` se abre desde un `editModal` (Corregir), JS setea explícitamente `z-index` del `editModal` padre a `1045` y el selector a `1070` más `backdrop:static`. Sin esto, clic en mini-mapa roba el foco, Corregir se cierra silenciosamente, y el usuario no puede aplicar corrección de zona.
- La función `abrirSelectorZona` (`index.html` JS) detecta si el input pertenece a un `editModal` y aplica estos z-index + backdrop static.

## Validación duplicados OT

- `/ingresar` y `/confirmar_importado` ejecutan cada uno un `SELECT ... WHERE orden_trabajo ILIKE ... AND hora_fin IS NULL` antes de insertar.
- Endpoint AJAX `/verificar_ot_duplicada` provee validación en tiempo real en formulario de registro.
- Cuando `confirmar_importado` retorna `duplicada:true`, el handler JSON frontend restaura el checkbox y mantiene la fila en la tabla preview (NO borra la fila pendiente de la sesión).

## Alias de zonas soportados por UI (SINONIMOS)

- Definidos una vez en `api_validar_zonas` (app.py:553+). Ejemplo: `ESTACION22 -> E22`, `CAJA TIPO 1 RAMAL D1 -> D1`.
- Nuevo alias: añadir a dict `SINONIMOS` y verificar con `POST /api/validar_zonas` (test client).
- Actualmente (~17 entradas). Si se detecta nuevo formato Excel, añadir su mapeo aquí.

## Historial de cambios recientes

- **Monitoreo persistente**: Cambio en la query de `/` (index) para mostrar TODOS los trabajos no archivados con estado 'En Vía' o 'Liberado', sin filtro de fecha. Así los trabajos finalizados permanecen visibles en la ventana de monitoreo hasta que el usuario pulsa "Archivar Turno".
- **Zonas peatonales en monitoreo**: Corregido template `index.html` para mostrar `ubicacion_zona_peatonal` (índice 7) junto a `ubicacion_zona` en la tabla de monitoreo y en el modal "Corregir". Arreglados índices de `num_personas` (8), `tetra` (11), `comentario` (21).
- **Selector zonas - todo visible**: Eliminado filtro por tipo bivial/peatonal en `renderMiniMapa()` y `renderChipsPorCategoria()` — ahora se ven **todas** las zonas (PTSA, TK1-TK7, D1, D2, NTA, NTP, etc.) siempre, sin importar el input que abre el selector.
- **NTA/NTP**: Agregadas a `ZONE_POSITIONS` (app.py:201-202) y clasificadas como `ptsa` en `_clasificar_zona`.
- **Filtro chips/popover por categoría**: `renderChipsPorCategoria` y `onClicAreaZona` ahora respetan `filtroMapa` — en modo "Estaciones" solo se ven/eligen estaciones, sin tramos mezclados.
- **Acceso a vía en importación**: Tanto en `importar_excel` como `importar_texto` y `_parsear_fila_fija`, ahora se salta filas solo si `acceso == 'NO'`; antes solo aceptaba `'SI'`. `estado` debe ser `'AUTORIZADA'` (antes era `'CONFIRMADA'`).
- **Respaldo BD**: Se sugirió `pg_dump` diario como backup de Neon. Tener clon local del repo para emergencias si Render cae.

## Migración BD requerida (Neon / PostgreSQL)

Para soportar zonas peatonales + biviales en el mismo permiso, ejecutar:

```sql
ALTER TABLE seguimiento_vias 
ADD COLUMN IF NOT EXISTS ubicacion_zona_peatonal TEXT;
```