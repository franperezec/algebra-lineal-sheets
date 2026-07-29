# -*- coding: utf-8 -*-
"""Tests de la capa de datos con pandas: cargar_datos, guardar_datos,
matriz_diseno."""
import numpy as np
import pytest

pd = pytest.importorskip('pandas')

from algebra_lineal.core import cargar_datos, guardar_datos, matriz_diseno


def _csv(tmp_path, nombre, contenido, encoding='utf-8'):
    ruta = tmp_path / nombre
    ruta.write_text(contenido, encoding=encoding)
    return str(ruta)


def test_cargar_csv_con_encabezados(tmp_path):
    ruta = _csv(tmp_path, 'd.csv', 'consumo,ingreso\n90.5,100\n98,110\n')
    datos = cargar_datos(ruta)
    assert list(datos.columns) == ['consumo', 'ingreso']
    assert len(datos) == 2
    assert datos['consumo'].iloc[0] == 90.5


def test_cargar_csv_regional(tmp_path):
    # separador ';' y decimales con coma (Excel en español)
    ruta = _csv(tmp_path, 'd.csv', 'consumo;ingreso\n90,5;100\n98;110\n',
                encoding='utf-8-sig')
    datos = cargar_datos(ruta)
    assert list(datos.columns) == ['consumo', 'ingreso']
    assert datos['consumo'].iloc[0] == 90.5


def test_cargar_csv_sin_encabezados(tmp_path):
    ruta = _csv(tmp_path, 'd.csv', '1,2\n3,4\n')
    datos = cargar_datos(ruta)
    assert list(datos.columns) == ['x1', 'x2']
    assert datos['x1'].tolist() == [1, 3]


def test_cargar_archivo_inexistente(tmp_path, capsys):
    assert cargar_datos(str(tmp_path / 'nada.csv')) is None
    assert 'No existe' in capsys.readouterr().out


def test_guardar_y_recargar(tmp_path):
    datos = pd.DataFrame({'a': [1.5, 2.5], 'b': [3.0, 4.0]})
    ruta = str(tmp_path / 'salida.csv')
    assert guardar_datos(datos, ruta) is True

    recargado = cargar_datos(ruta)
    assert list(recargado.columns) == ['a', 'b']
    assert recargado['a'].tolist() == [1.5, 2.5]


def test_matriz_diseno_basica():
    datos = pd.DataFrame({
        'consumo': [90.0, 98.0, 106.0],
        'ingreso': [100.0, 110.0, 120.0],
        'precio': [5.0, 6.0, 7.0],
    })
    X, y, nombres = matriz_diseno(datos, y='consumo',
                                  x=['ingreso', 'precio'])
    assert nombres == ['const', 'ingreso', 'precio']
    assert X.shape == (3, 3)
    assert y.shape == (3, 1)
    assert np.array_equal(X[:, 0], np.ones(3))
    assert np.array_equal(X[:, 1], datos['ingreso'].to_numpy())


def test_matriz_diseno_x_por_defecto_y_sin_constante():
    datos = pd.DataFrame({'y': [1.0, 2.0], 'a': [3.0, 4.0], 'b': [5.0, 6.0]})
    X, y, nombres = matriz_diseno(datos, y='y', constante=False)
    assert nombres == ['a', 'b']
    assert X.shape == (2, 2)


def test_matriz_diseno_columna_inexistente(capsys):
    datos = pd.DataFrame({'a': [1.0], 'b': [2.0]})
    X, y, nombres = matriz_diseno(datos, y='consumo')
    assert X is None
    salida = capsys.readouterr().out
    assert 'no encontradas' in salida
    assert 'a, b' in salida


def test_matriz_diseno_descarta_filas_faltantes(capsys):
    datos = pd.DataFrame({
        'y': [1.0, None, 3.0],
        'a': [4.0, 5.0, 6.0],
    })
    X, y, nombres = matriz_diseno(datos, y='y')
    assert X.shape == (2, 2)
    assert 'descartaron 1 filas' in capsys.readouterr().out


def test_matriz_diseno_columna_de_texto(capsys):
    datos = pd.DataFrame({'y': [1.0, 2.0], 'ciudad': ['Quito', 'Loja']})
    X, y, nombres = matriz_diseno(datos, y='y', x=['ciudad'])
    assert X is None
    assert 'sin datos numéricos' in capsys.readouterr().out
