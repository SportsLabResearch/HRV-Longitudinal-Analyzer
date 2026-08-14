"""Codificación local y publicación segura de las imágenes públicas."""

from __future__ import annotations

import json
import re
import shutil
import subprocess
from datetime import datetime
from pathlib import Path


IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".webp"}
PRIVATE_STATE_DIR = ".private"
MAPPING_FILE = "participant_codes.json"


def _participant_dirs(data_dir: Path) -> list[Path]:
    if not data_dir.is_dir():
        return []
    return sorted(
        (p for p in data_dir.iterdir() if p.is_dir() and not p.name.startswith(".")),
        key=lambda p: p.name.casefold(),
    )


def _load_mapping(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"No se pudo leer la correspondencia privada: {exc}") from exc
    mapping = payload.get("participants", {})
    if not isinstance(mapping, dict):
        raise RuntimeError("La correspondencia privada no tiene un formato válido.")
    return {str(name): str(code) for name, code in mapping.items()}


def _save_mapping(path: Path, mapping: dict[str, str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "warning": "ARCHIVO PRIVADO. NO SUBIR A GITHUB.",
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "participants": dict(sorted(mapping.items(), key=lambda item: item[1])),
    }
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)


def _next_code(mapping: dict[str, str]) -> str:
    used = []
    for code in mapping.values():
        if code.startswith("Sujeto_") and code[7:].isdigit():
            used.append(int(code[7:]))
    return f"Sujeto_{max(used, default=0) + 1:03d}"


def _leading_participant_number(name: str) -> int | None:
    """Obtiene el identificador numérico inicial de una carpeta privada."""

    match = re.match(r"^\s*0*(\d+)(?:\D|$)", str(name))
    return int(match.group(1)) if match else None


def _reconcile_renamed_participants(mapping: dict[str, str], participants: list[Path]) -> int:
    """Conserva el código cuando cambia el texto del nombre de una carpeta.

    Si la carpeta comienza por un número estable (por ejemplo, ``002_``), ese
    número permite reconocer a la misma participante aunque se elimine o añada
    texto al nombre de la carpeta.
    """

    renamed = 0
    current_names = {participant.name for participant in participants}
    historical_by_number: dict[int, list[tuple[str, str]]] = {}
    for old_name, code in mapping.items():
        number = _leading_participant_number(old_name)
        if number is not None:
            historical_by_number.setdefault(number, []).append((old_name, code))

    for participant in participants:
        if participant.name in mapping:
            continue
        number = _leading_participant_number(participant.name)
        candidates = historical_by_number.get(number, []) if number is not None else []
        candidate_codes = {code for _, code in candidates}
        if len(candidate_codes) != 1:
            continue
        code = next(iter(candidate_codes))
        mapping[participant.name] = code
        for old_name, old_code in candidates:
            if old_code == code and old_name not in current_names:
                mapping.pop(old_name, None)
        renamed += 1
    return renamed


def sync_public_data(project_root: str | Path) -> dict[str, int | str]:
    """Copia imágenes privadas a carpetas codificadas con códigos estables."""

    root = Path(project_root).resolve()
    data_dir = root / "Datos"
    public_dir = root / "Datos_publicos"
    mapping_path = root / PRIVATE_STATE_DIR / MAPPING_FILE

    data_dir.mkdir(parents=True, exist_ok=True)
    public_dir.mkdir(parents=True, exist_ok=True)
    participants = _participant_dirs(data_dir)

    mapping = _load_mapping(mapping_path)
    codes = list(mapping.values())
    if len(codes) != len(set(codes)):
        raise RuntimeError("La correspondencia privada contiene códigos duplicados.")

    renamed = _reconcile_renamed_participants(mapping, participants)
    new_codes = 0
    for participant in participants:
        if participant.name not in mapping:
            mapping[participant.name] = _next_code(mapping)
            new_codes += 1

    copied = 0
    for participant in participants:
        code = mapping[participant.name]
        destination = public_dir / code
        destination.mkdir(parents=True, exist_ok=True)
        for source in sorted(participant.rglob("*")):
            if not source.is_file() or source.suffix.lower() not in IMAGE_SUFFIXES:
                continue
            relative = source.relative_to(participant)
            target = destination / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            if not target.exists() or source.stat().st_size != target.stat().st_size or source.stat().st_mtime_ns > target.stat().st_mtime_ns:
                shutil.copy2(source, target)
                copied += 1

    _save_mapping(mapping_path, mapping)
    return {
        "participants": len(participants),
        "images": copied,
        "new_codes": new_codes,
        "renamed": renamed,
        "status": "ok",
    }


def _run_git(root: Path, *args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["git", *args], cwd=root, text=True, capture_output=True, encoding="utf-8", errors="replace"
    )


def publish_public_data_to_github(project_root: str | Path) -> tuple[bool, str]:
    """Publica solo Datos_publicos en el repositorio Git configurado."""

    root = Path(project_root).resolve()
    if not (root / ".git").exists():
        return False, "Esta carpeta no es un repositorio Git."

    summary = sync_public_data(root)
    public_dir = root / "Datos_publicos"
    if not any(p.is_file() for p in public_dir.rglob("*")):
        return False, "Datos_publicos no contiene imágenes para publicar."

    for private_path in ("Datos", ".private"):
        ignored_private = _run_git(root, "check-ignore", "-q", private_path)
        if ignored_private.returncode != 0:
            return False, f"Protección cancelada: {private_path} debe estar excluida por .gitignore."

    add = _run_git(root, "add", "-A", "--", "Datos_publicos")
    if add.returncode != 0:
        return False, add.stderr.strip() or "No se pudo preparar Datos_publicos."

    staged = _run_git(root, "diff", "--cached", "--name-only", "--", "Datos_publicos")
    if staged.returncode != 0:
        return False, staged.stderr.strip() or "No se pudo comprobar la actualización."
    if not staged.stdout.strip():
        return True, "Datos_publicos ya está actualizado; no hay cambios que subir."

    commit = _run_git(
        root,
        "commit",
        "-m",
        f"Actualizar datos públicos {datetime.now().strftime('%Y-%m-%d')}",
        "--",
        "Datos_publicos",
    )
    if commit.returncode != 0:
        return False, commit.stderr.strip() or commit.stdout.strip() or "No se pudo crear el commit."

    push = _run_git(root, "push")
    if push.returncode != 0:
        return False, push.stderr.strip() or push.stdout.strip() or "No se pudo actualizar GitHub."

    return True, "Datos_publicos se ha actualizado correctamente en GitHub."


def show_public_data_status(project_root: str | Path) -> None:
    root = Path(project_root).resolve()
    mapping = _load_mapping(root / PRIVATE_STATE_DIR / MAPPING_FILE)
    public_dir = root / "Datos_publicos"
    images = sum(1 for p in public_dir.rglob("*") if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES) if public_dir.exists() else 0
    print("\nDATOS PÚBLICOS CODIFICADOS")
    print(f"Participantes codificados: {len(mapping)}")
    print(f"Imágenes públicas:         {images}")
    print("Los nombres reales y su correspondencia permanecen únicamente en local.")
