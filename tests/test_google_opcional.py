# -*- coding: utf-8 -*-
"""Tests del soporte opcional de Google: el paquete debe funcionar sin gspread."""
import os
import pathlib
import subprocess
import sys

import pytest

import algebra_lineal
import algebra_lineal.core as core


@pytest.fixture
def sin_google(monkeypatch):
    """Simula un entorno sin las dependencias de Google instaladas."""
    for mod in ('gspread', 'gspread.utils', 'google.auth', 'google.colab'):
        monkeypatch.setitem(sys.modules, mod, None)


@pytest.fixture
def fuente_restaurada():
    """Restaura la fuente activa y el cliente de Google tras el test."""
    fuente, cliente = core.spreadsheet_name, core.gc
    yield
    core.spreadsheet_name, core.gc = fuente, cliente


def test_import_sin_google_en_subproceso():
    """`import algebra_lineal` debe funcionar solo con numpy + openpyxl."""
    codigo = (
        "import sys, importlib.abc\n"
        "class Bloqueador(importlib.abc.MetaPathFinder):\n"
        "    def find_spec(self, nombre, path=None, target=None):\n"
        "        if nombre.split('.')[0] in ('gspread', 'google'):\n"
        "            raise ImportError(nombre + ' bloqueado para el test')\n"
        "        return None\n"
        "sys.meta_path.insert(0, Bloqueador())\n"
        "import algebra_lineal\n"
        "print('IMPORT_OK')\n"
    )
    raiz_repo = pathlib.Path(__file__).resolve().parents[1]
    entorno = {**os.environ, 'PYTHONUTF8': '1'}  # emojis de la bienvenida en Windows
    resultado = subprocess.run(
        [sys.executable, '-c', codigo], capture_output=True, text=True,
        encoding='utf-8', cwd=str(raiz_repo), env=entorno)
    assert resultado.returncode == 0, resultado.stderr
    assert 'IMPORT_OK' in resultado.stdout


def test_requiere_google_lanza_importerror_con_extra(sin_google):
    with pytest.raises(ImportError, match=r'algebra-lineal-sheets\[google\]'):
        core._requiere_google()


def test_requiere_google_devuelve_modulos_si_instalado():
    gspread, value_render, default = core._requiere_google()
    assert hasattr(gspread, 'authorize')
    assert hasattr(value_render, 'unformatted')
    assert callable(default)


def test_configurar_google_sin_gspread_mensaje_espanol(
        sin_google, fuente_restaurada, capsys):
    assert core.configurar(sheet='matrices') is False
    salida = capsys.readouterr().out
    assert 'algebra-lineal-sheets[google]' in salida
    assert 'mi_archivo.xlsx' in salida


def test_configurar_excel_funciona_sin_google(
        sin_google, fuente_restaurada, tmp_path):
    ruta = str(tmp_path / 'm.xlsx')
    assert core.configurar(sheet=ruta) is True


def test_backend_google_sin_gspread_lanza_importerror(sin_google):
    with pytest.raises(ImportError, match=r'algebra-lineal-sheets\[google\]'):
        core._BackendGoogle('cualquier_sheet')


def test_verificar_instalacion_sin_google_devuelve_true(sin_google, capsys):
    assert algebra_lineal.verificar_instalacion() is True
    salida = capsys.readouterr().out
    assert 'opcional' in salida
    assert '❌' not in salida
