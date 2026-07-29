# -*- coding: utf-8 -*-
"""Tests del registro de fuentes de _conectar (elección de backend y mensajes)."""
import openpyxl
import pytest

import algebra_lineal.core as core
from algebra_lineal.core import _BackendCSV, _BackendExcel, _conectar


def test_conectar_elige_backend_por_tipo(tmp_path):
    ruta_xlsx = str(tmp_path / 'm.xlsx')
    libro = openpyxl.Workbook()
    libro.active.title = 'A'
    libro.save(ruta_xlsx)
    assert isinstance(_conectar(ruta_xlsx), _BackendExcel)

    carpeta = tmp_path / 'csvs'
    carpeta.mkdir()
    assert isinstance(_conectar(str(carpeta)), _BackendCSV)

    ruta_csv = tmp_path / 'a.csv'
    ruta_csv.write_text('1;2\n', encoding='utf-8')
    assert isinstance(_conectar(str(ruta_csv)), _BackendCSV)


def test_conectar_excel_inexistente_avisa(tmp_path, capsys):
    assert _conectar(str(tmp_path / 'no_existe.xlsx')) is None
    assert 'No existe el archivo Excel' in capsys.readouterr().out


def test_conectar_csv_inexistente_avisa(tmp_path, capsys):
    assert _conectar(str(tmp_path / 'no_existe.csv')) is None
    assert 'No existe la fuente CSV' in capsys.readouterr().out


def test_conectar_google_sin_configurar_avisa(monkeypatch, capsys):
    """Todo lo que no es ruta local cae al fallback de Google."""
    monkeypatch.setattr(core, 'gc', None)
    assert _conectar('un_sheet_cualquiera') is None
    assert 'Sistema no configurado' in capsys.readouterr().out


def test_conectar_para_escribir_crea_excel_nuevo(tmp_path):
    """para_escribir=True propaga crear_si_no_existe al backend."""
    backend = _conectar(str(tmp_path / 'nuevo.xlsx'), para_escribir=True)
    assert isinstance(backend, _BackendExcel)
