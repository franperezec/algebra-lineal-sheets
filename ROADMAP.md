# 🗺️ Roadmap de `algebra-lineal-sheets`

La ventaja diferenciadora de esta librería es ser un **puente hoja de cálculo ↔
Python para no-programadores, en español, bidireccional y sin configuración**.
Este roadmap ordena las mejoras futuras alrededor de esa ventaja.

## ✅ Fase 1 — Confiabilidad (v1.2.0, completada)

- Lectura sin pérdida: números nativos desde Google Sheets (`UNFORMATTED`) y
  openpyxl; los decimales con coma (`3,5` / `1.234,56`) se convierten
  correctamente en vez de volverse `0.0` en silencio.
- Validación con mensajes en español y coordenadas: texto en celdas
  (p. ej. filas de encabezados) y celdas vacías intermedias ahora dan un error
  claro en vez de corromper la matriz.
- Exportación siempre con números nativos (Excel/Sheets aplican el formato
  regional del usuario), rechazo de NaN/infinito con mensaje claro, aviso
  cuando el Excel contiene fórmulas, pestañas de Google dimensionadas según
  los datos y firma de `update` compatible con gspread 7.
- Primera suite de tests (`tests/`).

## 🇪🇸 Fase 2 — Español-first

- Jerarquía interna de errores traducida a mensajes en español en un solo
  punto (hoy los errores más informativos de numpy/gspread/openpyxl salen
  crudos en inglés).
- Distinguir causas hoy indistinguibles: sin internet, sin permisos, nombre
  mal escrito, cuota de la API de Google agotada.
- Corregir el mensaje engañoso "Variables no encontradas" cuando la variable
  existe pero su tipo no es exportable (listas, DataFrames...).
- Exportar `verificar_instalacion()` en `__all__` y que también revise
  openpyxl.

## 📴 Fase 3 — Modo offline de verdad

- Imports perezosos de gspread/google-auth: que `import algebra_lineal`
  funcione solo con numpy + openpyxl, y ofrecer el extra
  `pip install algebra-lineal-sheets[google]` para Colab.
- Backend CSV (carpeta de archivos .csv) mediante un registro de fuentes en
  `_conectar` en vez del if/elif actual.
- Rendimiento en Excel: una sola carga/guardado del libro por llamada a
  `exportar()` (hoy es una por variable).

## 🎓 Fase 4 — Courseware (valor agregado único)

- `verificar('x_sol')`: comparar la respuesta del estudiante contra la
  solución del profesor (pestaña oculta u otro archivo) con tolerancia
  numérica y explicación en español (por qué `0.30000000000000004 ≈ 0.3`).
- Confirmación o filtro en `exportar()` sin argumentos (hoy exporta cualquier
  variable numérica suelta, incluidos contadores de loops).
- Plantillas de ejercicios por pestaña y generación de un archivo por
  estudiante.
- **Prerrequisito técnico**: reemplazar la inyección de variables dependiente
  de la profundidad de pila (`f_back.f_back` en `core.py`) por un parámetro
  de namespace explícito; hoy impide construir funciones encima de
  `importar()`.

## 🔧 Transversal

- Ampliar la suite de tests (hoy cubre el pipeline de datos y los backends).
- Unificar los clasificadores de tipo duplicados (`_obtener_dimensiones_y_tipo`,
  `workspace`, `_convertir_a_valor`) y el predicado de variable exportable
  (duplicado en `exportar` y `listar_variables_exportables`).
- Eliminar código muerto (`MatrixLoader`) y leer la versión desde un solo
  lugar (flit puede tomarla de `__version__`).
