import json
from pathlib import Path

from modules.public_data import sync_public_data


def test_codigos_estables_y_copia_de_imagenes(tmp_path: Path):
    ana = tmp_path / "Datos" / "Ana López"
    beatriz = tmp_path / "Datos" / "Beatriz Pérez"
    ana.mkdir(parents=True)
    beatriz.mkdir(parents=True)
    (ana / "registro.jpg").write_bytes(b"imagen-1")
    (beatriz / "registro.png").write_bytes(b"imagen-2")

    sync_public_data(tmp_path)
    mapping_path = tmp_path / ".private" / "participant_codes.json"
    first = json.loads(mapping_path.read_text(encoding="utf-8"))["participants"]

    carla = tmp_path / "Datos" / "Carla Ruiz"
    carla.mkdir()
    (carla / "registro.jpeg").write_bytes(b"imagen-3")
    sync_public_data(tmp_path)
    second = json.loads(mapping_path.read_text(encoding="utf-8"))["participants"]

    assert first["Ana López"] == second["Ana López"] == "Sujeto_001"
    assert second["Carla Ruiz"] == "Sujeto_003"
    assert (tmp_path / "Datos_publicos" / "Sujeto_001" / "registro.jpg").exists()
    assert not (tmp_path / "Datos_publicos" / "Ana López").exists()
