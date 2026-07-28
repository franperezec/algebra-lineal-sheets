"""
📊 Ejemplo: modo Excel local (sin internet ni cuenta de Google)
===============================================================

Este ejemplo crea un archivo Excel con matrices, las importa,
opera con ellas y exporta los resultados al mismo archivo.

Ejecutar:  python ejemplo_excel_local.py
"""

import numpy as np
from algebra_lineal import *

# 1. Elegir un archivo Excel como fuente de trabajo
#    (no necesita existir: exportar() lo crea)
configurar(sheet='mis_matrices.xlsx')

# 2. Crear matrices de ejemplo y guardarlas en el Excel
A = np.array([[2.0, 1.0], [1.0, 3.0]])
b = np.array([[5.0], [10.0]])
exportar('A', 'b')

# 3. Ver qué contiene el archivo
workspace()

# 4. Importar y resolver el sistema Ax = b
importar('A', 'b')
x = np.linalg.solve(A, b)
print(f"\nSolución del sistema Ax = b:\n{x}")

# 5. Exportar el resultado — queda en la pestaña 'x' del mismo .xlsx
exportar('x')

print("\n✅ Abre 'mis_matrices.xlsx' en Excel para ver las pestañas A, b y x")
