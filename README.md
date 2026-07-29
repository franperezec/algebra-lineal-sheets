# 📚 Álgebra Lineal con Google Sheets y Excel
**Álgebra lineal simplificada para estudiantes con integración perfecta a Google Sheets y Excel.**

Permite a estudiantes y profesores trabajar con matrices almacenadas en Google Sheets (por nombre o por enlace) o en archivos Excel locales, usando Python de forma intuitiva y sencilla. Perfecto para cursos de álgebra lineal, análisis numérico y ciencias de datos.

## 🚀 Instalación

```bash
pip install algebra-lineal-sheets
```

¡Y listo! No necesitas configurar nada más.

### 🎓 Materiales de la clase (sin clonar nada)

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/franperezec/algebra-lineal-sheets/blob/main/ejemplos/numpy_para_economistas.ipynb)

- **Notebook de la clase** *"NumPy para economistas"*: ábrelo directo en Google
  Colab con el botón de arriba, o
  [descárgalo aquí](https://raw.githubusercontent.com/franperezec/algebra-lineal-sheets/main/ejemplos/numpy_para_economistas.ipynb)
  (clic derecho → *"Guardar enlace como..."*).
- **Excel de ejemplo** con las matrices de la clase (`A`, `B`, `v1`, `v2`, `b0`):
  [descargar matrices_ejemplo.xlsx](https://raw.githubusercontent.com/franperezec/algebra-lineal-sheets/main/ejemplos/matrices_ejemplo.xlsx).

El notebook funciona en dos modos:

- **En Colab**: usa Google Sheets como en clase (`configurar()`), o sube
  `matrices_ejemplo.xlsx` al entorno de Colab (icono 📁 → subir) y usa
  `configurar(sheet='matrices_ejemplo.xlsx')`.
- **En local (VS Code / Jupyter)**: guarda el notebook y el Excel en la misma
  carpeta y usa el modo Excel local — sin internet ni cuenta de Google.

### 💻 Trabajar en local (clonar el repositorio) — opcional

Clonar el repositorio **no es necesario** para usar la librería (basta el
`pip install` de arriba) ni para la clase (los materiales se descargan en la
sección anterior). Clona si quieres todo junto en una carpeta o vas a
contribuir al proyecto:

```bash
# 1. Clonar el repositorio
git clone https://github.com/franperezec/algebra-lineal-sheets.git
cd algebra-lineal-sheets

# 2. Crear y activar un entorno virtual (Python 3.8 o superior)
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # macOS / Linux

# 3. Instalar la librería y las herramientas de notebook
pip install -r requirements.txt
```

Esto instala `algebra-lineal-sheets` con todas sus dependencias (numpy,
gspread, google-auth, openpyxl), más `ipykernel`, `nbconvert`, `matplotlib`
y `pandas` para ejecutar el notebook de la clase
[`ejemplos/numpy_para_economistas.ipynb`](ejemplos/numpy_para_economistas.ipynb).
En VS Code, abre el notebook y selecciona el kernel del `.venv`.

> 🛠️ **Opcional — ¿vas a modificar el código de la librería?** Por defecto
> `requirements.txt` instala la versión publicada en PyPI, y esa es la forma
> recomendada para la clase. Si además quieres que tus cambios locales en
> `algebra_lineal/` se usen sin reinstalar, instala la versión local en modo
> editable con `pip install -e .`.

> 📦 **¿Falta alguna librería?** Si al ejecutar un ejemplo aparece
> `ModuleNotFoundError` (por ejemplo con `scipy` u otra librería que usemos
> a futuro), activa el entorno e instálala con pip:
>
> ```bash
> .venv\Scripts\activate            # Windows
> # source .venv/bin/activate      # macOS / Linux
> pip install scipy                 # o la librería que falte
> ```

**⌨️ Atajos de teclado útiles en VS Code:**

| Acción | Windows / Linux | macOS |
|---|---|---|
| Abrir la terminal integrada (para escribir los comandos de arriba) | `Ctrl` + `` ` `` | `Ctrl` + `` ` `` |
| Paleta de comandos | `Ctrl` + `Shift` + `P` | `Cmd` + `Shift` + `P` |

💡 Alternativa sin comandos: abre la paleta (`Ctrl/Cmd` + `Shift` + `P`),
escribe **"Python: Create Environment..."**, elige **Venv**, selecciona tu
Python y marca `requirements.txt` cuando pregunte qué dependencias instalar —
VS Code crea el `.venv` y lo instala todo por ti.

**🗑️ ¿Quieres eliminar el entorno?** Borra la carpeta `.venv` en cualquier
momento: no afecta al código del proyecto y puedes regenerarla cuando la
necesites repitiendo los pasos 2 y 3 (o con "Python: Create Environment...").

## 📋 Uso Básico

### 1. Preparar Google Sheet
- Crear Google Sheet llamado `matrices`
- Añadir pestañas con nombres: `A`, `B`, `v`, etc.
- Llenar con datos numéricos (sin texto ni fórmulas)

### 2. Usar en Python

```python
# Importar y configurar (una vez por sesión)
from algebra_lineal import *
configurar()

# Ver qué matrices tienes disponibles
workspace()

# Importar matrices específicas
importar('A', 'B', 'v')

# Realizar operaciones de álgebra lineal
C = A @ B                      # Multiplicación matricial
suma = A + B                   # Suma de matrices
Ainv = np.linalg.inv(A)       # Matriz inversa
det_A = np.linalg.det(A)      # Determinante

# Exportar resultados de vuelta a Google Sheets
exportar('C', 'suma', 'Ainv')
```

## 📊 Ejemplo Completo

```python
from algebra_lineal import *
import numpy as np

# Configurar conexión con Google Sheets
configurar()

# Ver workspace
workspace()
# 🏢 WORKSPACE: 'matrices'
# ===========================================================================
# #   NOMBRE               DIMENSIONES  TIPO           
# ---------------------------------------------------------------------------
# 1   A                    3×3          📋 Matriz      
# 2   B                    3×3          📋 Matriz      
# 3   v                    3×1          📉 Vector columna

# Importar matrices necesarias
importar('A', 'B', 'v')

# Resolver sistema de ecuaciones Ax = b
b = v  # Usar vector v como término independiente
x = np.linalg.solve(A, b)

# Verificar solución
verificacion = A @ x - b
error = np.linalg.norm(verificacion)

print(f"Solución: x = {x}")
print(f"Error: {error:.2e}")

# Exportar resultados
exportar('x', 'verificacion')
```

## 🔧 Funciones Disponibles

| Función | Descripción | Ejemplo |
|---------|-------------|---------|
| `configurar()` | Configuración inicial | `configurar(sheet='prope2026')`, `configurar(sheet='matrices.xlsx')` o `configurar(sheet='matrices/')` |
| `workspace()` | Ver matrices de la fuente activa | `workspace()` |
| `importar()` | Importar matrices | `importar('A', 'B')` |
| `exportar()` | Exportar resultados | `exportar('C')` |
| `cambiar_sheet()` | Cambiar de fuente | `cambiar_sheet('proyecto2')` |
| `cargar_datos()` | Datos con columnas nombradas (pandas) | `datos = cargar_datos('consumo.csv')` |
| `guardar_datos()` | Guardar un DataFrame | `guardar_datos(datos, 'salida.csv')` |
| `matriz_diseno()` | Matriz de diseño para regresión | `X, y, nombres = matriz_diseno(datos, y='consumo')` |
| `listar_variables_exportables()` | Ver qué se puede exportar | `listar_variables_exportables()` |
| `version()` | Información del paquete | `version()` |
| `ayuda()` | Ayuda completa | `ayuda()` |

## 📚 Para Estudiantes

### Google Colab (Recomendado)

```python
# 1. Instalar paquete
!pip install algebra-lineal-sheets

# 2. Importar y configurar
from algebra_lineal import *
configurar()

# 3. ¡Empezar a trabajar!
workspace()
importar('A', 'B')
resultado = A @ B
exportar('resultado')
```

### Operaciones Comunes

```python
# Después de importar matrices A, B, v
C = A @ B                          # Multiplicación matricial
suma = A + B                       # Suma
transpuesta = A.T                  # Transpuesta
inversa = np.linalg.inv(A)         # Inversa (si existe)
determinante = np.linalg.det(A)    # Determinante
autovalores = np.linalg.eigvals(A) # Autovalores
rango = np.linalg.matrix_rank(A)   # Rango
norma = np.linalg.norm(v)          # Norma de vector
```

## 👨‍🏫 Para Profesores

### Ventajas Pedagógicas

- **Enfoque en matemáticas**: Los estudiantes se concentran en álgebra lineal, no en programación
- **Datos modificables**: Cambiar valores en Google Sheets sin tocar código
- **Colaborativo**: Fácil compartir matrices entre estudiantes
- **Visual**: Ver resultados inmediatamente en Google Sheets
- **Escalable**: Funciona igual para 10 o 1000 estudiantes

### Configuración de Clase

1. **Crear plantilla**: Google Sheet con matrices ejemplo
2. **Compartir plantilla**: Estudiantes hacen copia
3. **Dar instrucciones simples**:
   ```python
   !pip install algebra-lineal-sheets
   from algebra_lineal import *
   configurar()
   ```

### Ejemplo de Ejercicio

```python
# Ejercicio: Transformaciones lineales
importar('T', 'v1', 'v2', 'v3')  # Matriz T y vectores

# Aplicar transformación
w1 = T @ v1
w2 = T @ v2  
w3 = T @ v3

# Analizar propiedades
det_T = np.linalg.det(T)
es_invertible = abs(det_T) > 1e-10

# Exportar análisis
exportar('w1', 'w2', 'w3', 'det_T')
```

## 🛠️ Configuración Avanzada

### Fuentes de datos: nombre, enlace, Excel local o carpeta CSV

`configurar(sheet=...)` y `cambiar_sheet(...)` aceptan **cuatro formas** y
detectan automáticamente cuál es:

```python
# 1. Google Sheet por NOMBRE
configurar(sheet='prope2026')

# 2. Google Sheet por ENLACE (copia el enlace de edición del navegador)
configurar(sheet='https://docs.google.com/spreadsheets/d/1vpg.../edit')

# 3. Excel LOCAL en tu PC — no necesita internet ni cuenta de Google
configurar(sheet='C:/Users/ana/Documents/matrices.xlsx')

# 4. Carpeta de archivos CSV — cada .csv es una "pestaña" (A.csv, b0.csv...)
configurar(sheet='C:/Users/ana/Documents/matrices/')
```

En el modo CSV la lectura tolera el formato regional: separador `;` o `,`
(el Excel en español guarda CSV con `;`) y decimales con coma. También se
acepta un `.csv` suelto como fuente de una sola matriz.

También puedes cambiar de fuente en cualquier momento, o usar otra solo
para una operación puntual:

```python
cambiar_sheet('prope2026')
importar('A', 'B')

importar('A', sheet_name='otro_archivo')       # solo esta vez
exportar('C', sheet_name='resultados.xlsx')    # crea el Excel si no existe
```

⚠️ **Importante:** `importar('prope2026')` NO abre el archivo prope2026 —
busca una *pestaña* llamada prope2026 dentro del archivo activo.
Para cambiar de archivo usa `cambiar_sheet('prope2026')`.

### Modo Excel local

Ideal para trabajar **sin internet y sin cuenta de Google** (en tu PC con
Anaconda, VS Code, etc.). La estructura es la misma: una pestaña = una matriz.

```python
from algebra_lineal import *
configurar(sheet='matrices.xlsx')   # sin autenticación
workspace()
importar('A', 'B')
C = A @ B
exportar('C')                       # se guarda en el mismo .xlsx
```

Notas del modo Excel:
- Solo archivos `.xlsx` (no `.xls` antiguo)
- Si el archivo tiene fórmulas, debe haberse guardado desde Excel al menos
  una vez para que Python pueda leer los valores calculados
- Además, al guardar desde Python (`exportar()`) los valores calculados de
  las fórmulas se pierden para la librería: vuelve a abrir y guardar el
  archivo en Excel para que `importar()` los lea de nuevo (la librería te
  avisa cuando detecta fórmulas)
- `exportar()` crea el archivo si no existe
- Puedes escribir los decimales con coma o con punto (`3,5` o `3.5`): la
  librería entiende ambos, y al exportar verás los números en el formato
  regional de tu equipo

#### ✏️ Editar el Excel dentro de VS Code (Office Viewer)

Si trabajas en VS Code no necesitas abrir Excel: la extensión **Office Viewer**
permite ver y editar el archivo `.xlsx` directamente en el editor.

1. **Instalar la extensión:** abre la vista de Extensiones (`Ctrl+Shift+X`),
   busca **"Office Viewer"** (`cweijan.vscode-office`) e instálala. O desde la
   terminal:
   ```bash
   code --install-extension cweijan.vscode-office
   ```
2. **Abrir el archivo:** haz clic derecho sobre el `.xlsx` en el explorador de
   VS Code → **"Open With..." / "Abrir con..."** → elige **Office Viewer**
   (puedes marcarla como editor por defecto para todos los `.xlsx`).
3. **Editar y guardar:** modifica las celdas y guarda con `Ctrl+S`. Escribe
   valores numéricos directos (no fórmulas — ver la nota anterior sobre
   fórmulas y valores calculados).

### 📈 Datos, CSV y regresiones (pandas)

Para trabajar con **datos con columnas nombradas** (por ejemplo la matriz de
diseño de una regresión), la librería incluye tres funciones que usan pandas
(`pip install pandas`; en Google Colab ya viene instalado):

```python
from algebra_lineal import *
import numpy as np

# 1. Cargar datos: detecta separador (',' o ';'), decimales con coma y
#    si la primera fila trae los nombres de las variables
datos = cargar_datos('datos_consumo.csv')
# ✅ datos_consumo.csv → 12 filas × 3 columnas: consumo, ingreso, precio

# 2. Matriz de diseño: y = consumo; X = [const, ingreso, precio]
X, y, nombres = matriz_diseno(datos, y='consumo', x=['ingreso', 'precio'])

# 3. Regresión OLS con álgebra matricial pura
beta = np.linalg.solve(X.T @ X, X.T @ y)
for nombre, valor in zip(nombres, beta.flatten()):
    print(nombre, round(valor, 3))

# 4. Guardar un DataFrame en .csv o .xlsx
guardar_datos(datos, 'copia.xlsx')
```

Notas:
- Si la primera fila del archivo tiene texto, se toma como **nombres de las
  variables**; si es toda numérica, las columnas se llaman `x1, x2, ...`
- `matriz_diseno` descarta (avisando) las filas con datos faltantes y añade
  la columna de constante al inicio (desactivable con `constante=False`).
- Ejemplo completo en
  [`ejemplos/ejemplo_regresion.py`](ejemplos/ejemplo_regresion.py) con
  [`ejemplos/datos_consumo.csv`](ejemplos/datos_consumo.csv); al ejecutarlo
  genera `mi_regresion.xlsx`, que puedes comparar con el resultado esperado
  [`ejemplos/resultados_regresion.xlsx`](ejemplos/resultados_regresion.xlsx)
  (pestaña `beta`).

### Verificar Variables

```python
# Ver qué variables están disponibles para exportar
listar_variables_exportables()
```

## ❓ Solución de Problemas

### Error: "No se pudo abrir 'matrices'"
- ✅ Verificar que el Google Sheet existe
- ✅ Verificar que se llama exactamente 'matrices'
- ✅ Verificar permisos de acceso

### Importa las matrices de OTRO archivo (no el que quiero)
- ✅ El paquete siempre usa el archivo activo (por defecto `matrices`)
- ✅ Cambiar con `cambiar_sheet('mi_archivo')` o `configurar(sheet='mi_archivo')`
- ✅ Ver qué archivo está activo: aparece en los mensajes de `importar()` y `workspace()`

### Error: "Variable no encontrada"
- ✅ Ejecutar `importar()` antes de usar variables
- ✅ Verificar nombres exactos con `workspace()`

### Error de autenticación
- ✅ Ejecutar `configurar()` nuevamente
- ✅ En Colab: Runtime → Restart and run all

## 🔄 Actualización

```bash
pip install --upgrade algebra-lineal-sheets
```

## 📦 Requisitos

- Python 3.8+
- numpy >= 1.20.0
- gspread >= 5.0.0
- google-auth >= 2.0.0
- openpyxl >= 3.0.0 (para el modo Excel local)

Se instalan automáticamente con el paquete.

## 📄 Licencia

MIT License - Ver [LICENSE](https://github.com/franperezec/algebra-lineal-sheets/blob/main/LICENSE) para más detalles.

## 🤝 Contribuir

¡Las contribuciones son bienvenidas! Revisa el [ROADMAP](ROADMAP.md) para ver
las mejoras planificadas. Los tests se ejecutan con `pytest tests/`.

## 📧 Contacto

- **Autor:** Francisco Pérez Mogollón
- **Email:** francisco.perezxxi@gmail.com
- **PyPI:** https://pypi.org/project/algebra-lineal-sheets/

## 🔗 Enlaces Útiles

- [Google Colab](https://colab.research.google.com/)
- [Google Sheets](https://sheets.google.com/)
- [NumPy Documentation](https://numpy.org/doc/)

---

⭐ **¡Si te resulta útil, compártelo con otros profesores!** ⭐