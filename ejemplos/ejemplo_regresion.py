# -*- coding: utf-8 -*-
"""
📈 Ejemplo: regresión OLS con matrices (modo local, sin Google)

Carga un CSV con columnas nombradas, arma la matriz de diseño y estima
los coeficientes con álgebra matricial pura:

    beta = (X'X)^-1 X'y

Los datos (datos_consumo.csv) se generaron con
consumo = 10 + 0.8*ingreso - 1.5*precio + ruido pequeño,
así que los coeficientes estimados deben quedar cerca de esos valores.

Uso:
    python ejemplo_regresion.py
"""
import numpy as np
from algebra_lineal import *

# 1. Cargar los datos: detecta separador, decimales y encabezados
datos = cargar_datos('datos_consumo.csv')

# 2. Matriz de diseño: y = consumo; X = [const, ingreso, precio]
X, y, nombres = matriz_diseno(datos, y='consumo')

# 3. OLS matricial: beta = (X'X)^-1 X'y
beta = np.linalg.solve(X.T @ X, X.T @ y)

print("\nCoeficientes estimados:")
for nombre, valor in zip(nombres, beta.flatten()):
    print(f"   {nombre:>8}: {valor:8.3f}")

# 4. Bondad de ajuste rápida: R²
residuos = y - X @ beta
suma_residuos = (residuos.T @ residuos).item()
suma_total = ((y - y.mean()).T @ (y - y.mean())).item()
r2 = 1 - suma_residuos / suma_total
print(f"\nR² = {r2:.4f}")

# 5. Guardar el resultado en un Excel aparte (tu propia copia;
#    compárala con resultados_regresion.xlsx, el resultado esperado del repo)
exportar('beta', sheet_name='mi_regresion.xlsx')
print("\n✅ Revisa 'mi_regresion.xlsx' y compáralo con 'resultados_regresion.xlsx'")
