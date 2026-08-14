from pathlib import Path
from modules.auditoria import auditar_proyecto, guardar_resultados

ROOT = Path(__file__).resolve().parents[1]


def test_auditoria_devuelve_datos():
    data = auditar_proyecto()
    assert data["proyecto"] == "HRV-Longitudinal-Analyzer"
    assert "porcentaje" in data
    assert data["porcentaje"] >= 90


def test_auditoria_genera_resultados():
    data = auditar_proyecto()
    salida = guardar_resultados(data)
    assert salida.exists()
    assert (salida / "Resumen_Auditoria.txt").exists()
    assert list(salida.glob("*.docx"))
    assert list(salida.glob("*.xlsx"))
