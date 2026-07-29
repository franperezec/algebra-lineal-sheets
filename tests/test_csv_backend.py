# -*- coding: utf-8 -*-
"""Tests del backend CSV: carpeta como libro y archivo suelto."""
import numpy as np

import algebra_lineal.core as core
from algebra_lineal.core import _BackendCSV, _limpiar_y_validar


def test_carpeta_round_trip(tmp_path):
    carpeta = str(tmp_path)
    backend = _BackendCSV(carpeta)
    assert backend.nombres_pestanas() == []

    assert backend.escribir('A', [[1.5, 2.0], [3.0, 4.0]]) == 'creada'
    assert backend.escribir('b0', [[5.0], [6.0]]) == 'creada'
    assert backend.nombres_pestanas() == ['A', 'b0']

    assert _limpiar_y_validar('A', backend.leer('A')) == [[1.5, 2.0], [3.0, 4.0]]
    assert _limpiar_y_validar('b0', backend.leer('b0')) == [[5.0], [6.0]]

    assert backend.escribir('A', [[9.0]]) == 'actualizada'
    assert backend.escribir('A', [[7.0]], sobrescribir=False) is None
    assert _limpiar_y_validar('A', backend.leer('A')) == [[9.0]]


def test_csv_regional_punto_y_coma_con_comas(tmp_path):
    # Regresión regional: Excel en español guarda ';' como separador y ',' decimal
    ruta = tmp_path / 'M.csv'
    ruta.write_text('3,5;1,25\n2;4\n', encoding='utf-8-sig')

    backend = _BackendCSV(str(tmp_path))
    assert _limpiar_y_validar('M', backend.leer('M')) == [[3.5, 1.25], [2.0, 4.0]]


def test_archivo_suelto_como_fuente(tmp_path):
    ruta = tmp_path / 'matriz.csv'
    ruta.write_text('1,2\n3,4\n', encoding='utf-8')

    backend = _BackendCSV(str(ruta))
    assert backend.nombres_pestanas() == ['matriz']
    assert _limpiar_y_validar('matriz', backend.leer('matriz')) == \
        [[1.0, 2.0], [3.0, 4.0]]


def test_archivo_vacio(tmp_path):
    ruta = tmp_path / 'vacio.csv'
    ruta.write_text('', encoding='utf-8')
    backend = _BackendCSV(str(tmp_path))
    assert backend.leer('vacio') == []


def test_api_publica_con_carpeta(tmp_path):
    carpeta = str(tmp_path) + '\\'
    globals()['M_CSV_RT'] = np.array([[1.5, -2.5], [0.25, 4.0]])
    try:
        exportadas = core.exportar('M_CSV_RT', sheet_name=carpeta)
        assert exportadas == ['M_CSV_RT']
        assert (tmp_path / 'M_CSV_RT.csv').exists()

        resultado = core.importar('M_CSV_RT', sheet_name=carpeta)
        assert np.array_equal(resultado['M_CSV_RT'], globals()['M_CSV_RT'])
    finally:
        globals().pop('M_CSV_RT', None)
