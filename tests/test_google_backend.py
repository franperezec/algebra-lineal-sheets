# -*- coding: utf-8 -*-
"""Tests del backend Google con un worksheet falso (sin red ni credenciales)."""
from algebra_lineal.core import _BackendGoogle


class FakeWorksheet:
    def __init__(self):
        self.row_count = 100
        self.col_count = 26
        self.update_kwargs = None
        self.resized = None

    def clear(self):
        pass

    def resize(self, rows=None, cols=None):
        self.row_count, self.col_count = rows, cols
        self.resized = (rows, cols)

    def update(self, *args, **kwargs):
        # gspread 7 elimina la firma posicional vieja update('A1', datos)
        assert not args, "update() debe llamarse solo con keywords"
        self.update_kwargs = kwargs


def _backend_con(fake):
    backend = object.__new__(_BackendGoogle)
    backend._pestanas = {'A': fake}
    backend.descripcion = 'fake'
    return backend


def test_update_usa_keywords_compatibles_con_gspread7():
    fake = FakeWorksheet()
    accion = _backend_con(fake).escribir('A', [[1.0, 2.0]])
    assert accion == 'actualizada'
    assert fake.update_kwargs == {'values': [[1.0, 2.0]], 'range_name': 'A1'}


def test_redimensiona_si_los_datos_exceden_la_grilla():
    fake = FakeWorksheet()
    fake.row_count = 2
    fake.col_count = 2
    datos = [[float(j) for j in range(5)] for _ in range(4)]
    _backend_con(fake).escribir('A', datos)
    assert fake.resized == (4, 5)
