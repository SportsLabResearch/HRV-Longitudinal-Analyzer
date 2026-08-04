from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_archivos_principales_existen():
    requeridos = ["main.py", "config.py", "README.md", "CHANGELOG.md"]
    faltantes = [p for p in requeridos if not (ROOT / p).exists()]
    assert not faltantes, f"Faltan archivos principales: {faltantes}"


def test_carpetas_principales_existen():
    requeridas = ["modules", "Validation", "docs"]
    faltantes = [p for p in requeridas if not (ROOT / p).exists()]
    assert not faltantes, f"Faltan carpetas principales: {faltantes}"


def test_release_existe():
    releases = list(ROOT.glob("RELEASE*.md"))
    assert releases, "No hay archivos RELEASE documentados"
