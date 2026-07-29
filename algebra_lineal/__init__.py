"""
📚 ALGEBRA LINEAL - Paquete para Google Sheets y Excel
======================================================

Álgebra lineal simplificada con Google Sheets o Excel local.
Perfecto para estudiantes y profesores.

Instalación:
    pip install algebra-lineal-sheets
    pip install "algebra-lineal-sheets[google]"   # con Google Sheets (Colab ya lo trae)

Uso básico:
    from algebra_lineal import *
    configurar(sheet='prope2026')            # Google Sheet por nombre
    # configurar(sheet='https://docs...')    # o por enlace
    # configurar(sheet='matrices.xlsx')      # o Excel local (sin Google)
    # configurar(sheet='matrices/')          # o carpeta CSV local (sin Google)
    importar('A', 'B')
    C = A @ B
    exportar('C')

Autor: Francisco Pérez Mogollón
Email: francisco.perezxxi@gmail.com
Versión: 1.4.0
"""

# Importar todas las funciones principales del módulo core
from .core import (
    configurar,
    workspace,
    importar,
    exportar,
    cambiar_sheet,
    ayuda,
    version,
    listar_variables_exportables,
    cargar_datos,
    guardar_datos,
    matriz_diseno
)

# Definir qué se importa cuando alguien hace "from algebra_lineal import *"
__all__ = [
    'configurar',
    'workspace',
    'importar',
    'exportar',
    'cambiar_sheet',
    'ayuda',
    'version',
    'listar_variables_exportables',
    'cargar_datos',
    'guardar_datos',
    'matriz_diseno'
]

# Metadatos del paquete (IMPORTANTE: mantener sincronizado con pyproject.toml)
__version__ = "1.4.0"
__author__ = "Francisco Pérez Mogollón"
__email__ = "francisco.perezxxi@gmail.com"
__description__ = "Álgebra lineal simplificada con Google Sheets y Excel"
__url__ = "https://pypi.org/project/algebra-lineal-sheets/"

# Mensaje de bienvenida al importar (más profesional para PyPI)
print(f"📚 ALGEBRA LINEAL v{__version__}")
print("🎓 Para estudiantes de álgebra lineal")
print("🔧 Ejecuta: configurar() para empezar")
print("📖 Ayuda completa: ayuda()")
print("🔗 PyPI: https://pypi.org/project/algebra-lineal-sheets/")

# Función de conveniencia para verificar instalación


def verificar_instalacion():
    """
    Verifica la instalación: núcleo (numpy, openpyxl) y opcionales
    (Google Sheets, pandas).

    Devuelve True si el núcleo está completo; los opcionales solo se informan.
    """
    def _disponible(modulo):
        try:
            __import__(modulo)
            return True
        except ImportError:
            return False

    nucleo_ok = True
    for modulo in ("numpy", "openpyxl"):
        if _disponible(modulo):
            print(f"✅ {modulo}")
        else:
            print(f"❌ Falta {modulo} (dependencia del núcleo)")
            nucleo_ok = False

    if _disponible("gspread") and _disponible("google.auth"):
        print("✅ Google Sheets (gspread + google-auth)")
    else:
        print("ℹ️  Google Sheets no disponible (opcional)")
        print('   Para activarlo: pip install "algebra-lineal-sheets[google]"')
        print("   En Google Colab ya viene instalado")

    if _disponible("pandas"):
        print("✅ pandas (cargar_datos, matriz_diseno)")
    else:
        print("ℹ️  pandas no disponible (opcional): pip install pandas")

    if nucleo_ok:
        print("🚀 ¡Listo para usar algebra_lineal!")
    else:
        print("💡 Ejecuta: pip install --upgrade algebra-lineal-sheets")
    return nucleo_ok
