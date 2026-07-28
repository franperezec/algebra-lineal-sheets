# -*- coding: utf-8 -*-
"""Tests del backend Excel local y del round-trip por la API pública."""
import numpy as np
import openpyxl

import algebra_lineal.core as core
from algebra_lineal.core import _BackendExcel, _limpiar_y_validar


def test_escribir_celdas_numericas_no_texto(tmp_path):
    """Las celdas exportadas son números nativos: Excel les aplica el formato
    regional del usuario (coma o punto), nunca se escriben como texto."""
    ruta = str(tmp_path / 'm.xlsx')
    backend = _BackendExcel(ruta, crear_si_no_existe=True)
    backend.escribir('A', [[1.5, 2.0], [3.25, 4.0]])

    libro = openpyxl.load_workbook(ruta)
    celdas = [c.value for fila in libro['A'].iter_rows() for c in fila]
    assert all(isinstance(v, (int, float)) for v in celdas)
    assert celdas == [1.5, 2.0, 3.25, 4.0]


def test_round_trip_con_decimales(tmp_path):
    # Regresión: antes leer() pasaba todo por str() y reparseaba
    ruta = str(tmp_path / 'm.xlsx')
    backend = _BackendExcel(ruta, crear_si_no_existe=True)
    original = [[1.5, -2.75], [0.001, 1234.56]]
    backend.escribir('A', original)

    leidos = _limpiar_y_validar('A', backend.leer('A'))
    assert leidos == original


def test_sobrescribir_false_no_toca_la_pestana(tmp_path):
    ruta = str(tmp_path / 'm.xlsx')
    backend = _BackendExcel(ruta, crear_si_no_existe=True)
    backend.escribir('A', [[1.0]])
    assert backend.escribir('A', [[2.0]], sobrescribir=False) is None
    assert _limpiar_y_validar('A', backend.leer('A')) == [[1.0]]


def test_aviso_de_formulas_al_escribir(tmp_path, capsys):
    ruta = str(tmp_path / 'm.xlsx')
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = 'datos'
    hoja['A1'] = 1
    hoja['A2'] = '=A1*2'
    libro.save(ruta)

    backend = _BackendExcel(ruta)
    backend.escribir('B', [[9.0]])
    salida = capsys.readouterr().out
    assert 'FÓRMULAS' in salida


def test_exportar_importar_api_publica(tmp_path):
    """Round-trip completo por la API pública con fuente puntual."""
    ruta = str(tmp_path / 'pub.xlsx')
    globals()['M_TEST_RT'] = np.array([[1.5, 2.5], [3.5, 4.5]])
    try:
        exportadas = core.exportar('M_TEST_RT', sheet_name=ruta)
        assert exportadas == ['M_TEST_RT']

        resultado = core.importar('M_TEST_RT', sheet_name=ruta)
        assert np.array_equal(resultado['M_TEST_RT'], globals()['M_TEST_RT'])
    finally:
        globals().pop('M_TEST_RT', None)


def test_exportar_rechaza_nan(tmp_path, capsys):
    ruta = str(tmp_path / 'nan.xlsx')
    globals()['M_TEST_NAN'] = np.array([[1.0, np.nan]])
    try:
        exportadas = core.exportar('M_TEST_NAN', sheet_name=ruta)
        assert exportadas == []
        assert 'NaN' in capsys.readouterr().out
    finally:
        globals().pop('M_TEST_NAN', None)
