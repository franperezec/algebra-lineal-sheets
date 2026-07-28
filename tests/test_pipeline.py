# -*- coding: utf-8 -*-
"""Tests del pipeline puro: normalización, limpieza/validación y conversión."""
import numpy as np
import pytest

from algebra_lineal.core import (
    ErrorDeDatos,
    _contiene_no_finitos,
    _convertir_a_valor,
    _coordenada_celda,
    _limpiar_y_validar,
    _normalizar_celda,
    _preparar_datos_para_sheets,
)


class TestNormalizarCelda:
    def test_numeros_nativos(self):
        assert _normalizar_celda(3) == 3.0
        assert _normalizar_celda(3.5) == 3.5
        assert _normalizar_celda(np.float64(2.5)) == 2.5

    def test_punto_decimal(self):
        assert _normalizar_celda('3.5') == 3.5

    def test_coma_decimal(self):
        assert _normalizar_celda('3,5') == 3.5
        assert _normalizar_celda(' -2,75 ') == -2.75

    def test_miles_con_punto_y_coma_decimal(self):
        assert _normalizar_celda('1.234,56') == 1234.56

    def test_vacias(self):
        assert _normalizar_celda(None) is None
        assert _normalizar_celda('') is None
        assert _normalizar_celda('   ') is None

    def test_texto_se_devuelve_tal_cual(self):
        assert _normalizar_celda('hola') == 'hola'


class TestCoordenadaCelda:
    def test_coordenadas(self):
        assert _coordenada_celda(0, 0) == 'A1'
        assert _coordenada_celda(2, 1) == 'B3'
        assert _coordenada_celda(0, 26) == 'AA1'


class TestLimpiarYValidar:
    def test_matriz_normal(self):
        datos = [[1, 2], [3, 4]]
        assert _limpiar_y_validar('A', datos) == [[1.0, 2.0], [3.0, 4.0]]

    def test_coma_decimal_como_texto(self):
        # Regresión: antes '3,5' se convertía silenciosamente en 0.0
        datos = [['3,5', '1,25'], ['2', '4']]
        assert _limpiar_y_validar('A', datos) == [[3.5, 1.25], [2.0, 4.0]]

    def test_pestana_vacia(self):
        assert _limpiar_y_validar('A', []) == []
        assert _limpiar_y_validar('A', [[None, None], ['', '  ']]) == []

    def test_filas_y_columnas_vacias_colapsan(self):
        datos = [
            [None, None, None],
            [1, None, 2],
            [3, None, 4],
        ]
        assert _limpiar_y_validar('A', datos) == [[1.0, 2.0], [3.0, 4.0]]

    def test_texto_da_error_con_coordenada(self):
        # Regresión: antes 'hola' se convertía silenciosamente en 0.0
        datos = [[1, 'hola'], [3, 4]]
        with pytest.raises(ErrorDeDatos) as exc:
            _limpiar_y_validar('A', datos)
        mensaje = str(exc.value)
        assert 'B1' in mensaje
        assert 'hola' in mensaje
        assert "'A'" in mensaje

    def test_fila_de_encabezados_sugerencia(self):
        datos = [['x1', 'x2'], [1, 2], [3, 4]]
        with pytest.raises(ErrorDeDatos) as exc:
            _limpiar_y_validar('A', datos)
        assert 'ENCABEZADOS' in str(exc.value)

    def test_hueco_intermedio_da_error_con_coordenada(self):
        datos = [[1, None, 3], [4, 5, 6]]
        with pytest.raises(ErrorDeDatos) as exc:
            _limpiar_y_validar('A', datos)
        assert 'B1' in str(exc.value)

    def test_huecos_que_antes_desplazaban_en_silencio(self):
        # Regresión: antes esto importaba [[1,3,9],[4,5,6]] sin ningún aviso
        datos = [[1, None, 3, 9], [4, 5, 6, None]]
        with pytest.raises(ErrorDeDatos):
            _limpiar_y_validar('A', datos)


class TestConvertirAValor:
    def test_escalar(self):
        valor, tipo = _convertir_a_valor([[5.0]])
        assert valor == 5.0
        assert tipo == 'escalar'

    def test_vector_fila(self):
        valor, tipo = _convertir_a_valor([[1.0, 2.0, 3.0]])
        assert tipo == 'vector fila'
        assert valor.shape == (1, 3)

    def test_vector_columna(self):
        valor, tipo = _convertir_a_valor([[1.0], [2.0]])
        assert tipo == 'vector columna'
        assert valor.shape == (2, 1)

    def test_matriz(self):
        valor, tipo = _convertir_a_valor([[1.0, 2.0], [3.0, 4.0]])
        assert tipo == 'matriz'
        assert np.array_equal(valor, np.array([[1.0, 2.0], [3.0, 4.0]]))


class TestGuardasDeExportacion:
    def test_no_finitos_detectados(self):
        assert _contiene_no_finitos([[1.0, float('nan')]])
        assert _contiene_no_finitos([[float('inf')], [2.0]])
        assert not _contiene_no_finitos([[1.0, 2.0]])

    def test_preparar_formas(self):
        assert _preparar_datos_para_sheets(5) == [[5.0]]
        assert _preparar_datos_para_sheets(np.array([1.0, 2.0])) == [[1.0], [2.0]]
        matriz = _preparar_datos_para_sheets(np.array([[1, 2], [3, 4]]))
        assert matriz == [[1.0, 2.0], [3.0, 4.0]]
