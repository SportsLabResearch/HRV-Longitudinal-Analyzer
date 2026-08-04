import py_compile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def compilar(ruta: Path):
    py_compile.compile(str(ruta), doraise=True)


def test_main_compila():
    compilar(ROOT / "main.py")


def test_config_compila():
    compilar(ROOT / "config.py")


def test_modulos_principales_compilan():
    for archivo in (ROOT / "modules").glob("*.py"):
        compilar(archivo)
