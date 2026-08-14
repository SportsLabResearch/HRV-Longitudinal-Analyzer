# -*- coding: utf-8 -*-
"""
Preparador de distribución oficial.
Crea una carpeta organizada para publicar una release:
- Codigo_Fuente: ZIP del código fuente limpio.
- Windows: ZIP del ejecutable si ya existe.
- Release: documentación de la versión.
"""
from __future__ import annotations

import os
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

PROJECT_NAME = "HRV-Longitudinal-Analyzer"
VERSION = "v2.1.0"

EXCLUDE_DIRS = {
    ".git", ".github", "__pycache__", ".pytest_cache", ".mypy_cache",
    ".build_env_hrv", "build", "dist", "Distribucion", "Distribucion_Release",
    "Resultados", "Resultados_Auditoria", "Datos", "Datos_publicos", ".private",
    ".venv", "venv", "env",
}
EXCLUDE_SUFFIXES = {".pyc", ".pyo", ".tmp", ".log"}
EXCLUDE_NAMES = {
    "desktop.ini", "Thumbs.db", "participant_codes.json",
}


def _safe_version(version: str) -> str:
    return str(version).replace("/", "-").replace("\\", "-").strip()


def _should_exclude(path: Path, root: Path) -> bool:
    try:
        rel = path.relative_to(root)
    except Exception:
        rel = path
    parts = set(rel.parts)
    if parts & EXCLUDE_DIRS:
        return True
    if path.name in EXCLUDE_NAMES:
        return True
    if path.suffix.lower() in EXCLUDE_SUFFIXES:
        return True
    if path.name.lower().endswith(".zip"):
        return True
    if path.name.lower().endswith(".exe"):
        return True
    if "backup" in path.name.lower() and path.suffix.lower() == ".py":
        return True
    return False


def _zip_dir(source_dir: Path, zip_path: Path) -> None:
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for p in source_dir.rglob("*"):
            if p.is_file():
                zf.write(p, p.relative_to(source_dir))


def _zip_source(root: Path, zip_path: Path) -> int:
    count = 0
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for p in root.rglob("*"):
            if p.is_file() and not _should_exclude(p, root):
                zf.write(p, Path(PROJECT_NAME) / p.relative_to(root))
                count += 1
    return count


def _find_windows_exe(root: Path) -> Path | None:
    candidates = [
        root / "Distribucion" / "Windows" / PROJECT_NAME / f"{PROJECT_NAME}.exe",
        root / "dist" / PROJECT_NAME / f"{PROJECT_NAME}.exe",
        root / "Windows" / PROJECT_NAME / f"{PROJECT_NAME}.exe",
        root / f"{PROJECT_NAME}.exe",
    ]
    for p in candidates:
        if p.exists():
            return p
    for p in root.rglob(f"{PROJECT_NAME}.exe"):
        if "build" not in p.parts and ".build_env_hrv" not in p.parts:
            return p
    return None


def _copy_release_docs(root: Path, release_dir: Path) -> None:
    release_dir.mkdir(parents=True, exist_ok=True)
    docs = [
        "README.md", "LICENSE", "CHANGELOG.md", "CITATION.cff",
        "RELEASE_v2.1.0.md", "RELEASE_v1.5.0-a.md",
        "README_BUILD_LIMPIO.md", "README_DISTRIBUCION.md",
    ]
    for name in docs:
        src = root / name
        if src.exists():
            shutil.copy2(src, release_dir / name)


def preparar_distribucion(version: str = VERSION, root_dir: str | Path | None = None) -> Path:
    root = Path(root_dir or Path.cwd()).resolve()
    version_safe = _safe_version(version)
    base = root / "Distribucion_Release" / version_safe
    codigo_dir = base / "Codigo_Fuente"
    windows_dir = base / "Windows"
    release_dir = base / "Release"
    for d in (codigo_dir, windows_dir, release_dir):
        d.mkdir(parents=True, exist_ok=True)

    source_zip = codigo_dir / f"{PROJECT_NAME}_{version_safe}_SOURCE.zip"
    n_files = _zip_source(root, source_zip)

    _copy_release_docs(root, release_dir)

    exe = _find_windows_exe(root)
    windows_zip = windows_dir / f"{PROJECT_NAME}_Windows.zip"
    if exe:
        temp = windows_dir / PROJECT_NAME
        if temp.exists():
            shutil.rmtree(temp)
        # Si el exe pertenece a una carpeta onedir, copiamos la carpeta completa.
        if exe.parent.name == PROJECT_NAME:
            shutil.copytree(exe.parent, temp)
        else:
            temp.mkdir(parents=True, exist_ok=True)
            shutil.copy2(exe, temp / exe.name)
        for name in ("README.md", "LICENSE", "CHANGELOG.md"):
            src = root / name
            if src.exists():
                shutil.copy2(src, temp / name)
        _zip_dir(temp, windows_zip)
    else:
        windows_zip = None

    summary = base / "RESUMEN_DISTRIBUCION.txt"
    with summary.open("w", encoding="utf-8") as f:
        f.write("DISTRIBUCIÓN OFICIAL\n")
        f.write("=" * 60 + "\n")
        f.write(f"Proyecto: {PROJECT_NAME}\n")
        f.write(f"Versión : {version_safe}\n")
        f.write(f"Fecha   : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Código fuente: {source_zip}\n")
        f.write(f"Archivos fuente incluidos: {n_files}\n")
        if windows_zip:
            f.write(f"Windows ZIP: {windows_zip}\n")
        else:
            f.write("Windows ZIP: PENDIENTE. Ejecuta primero GENERAR_EXE_WINDOWS.bat\n")
        f.write("\nPara GitHub Release: subir solo HRV-Longitudinal-Analyzer_Windows.zip.\n")
        f.write("GitHub genera automáticamente Source code (zip) y Source code (tar.gz).\n")
    return base


def mostrar_preparar_distribucion(version: str = VERSION, root_dir: str | Path | None = None) -> None:
    print("\n" + "=" * 60)
    print(" PREPARAR DISTRIBUCIÓN OFICIAL")
    print("=" * 60)
    root = Path(root_dir or Path.cwd()).resolve()
    base = preparar_distribucion(version=version, root_dir=root)
    print(f"Proyecto : {PROJECT_NAME}")
    print(f"Versión  : {version}")
    print(f"Carpeta  : {base}")
    print("\nGenerado:")
    print("- Codigo_Fuente")
    print("- Windows")
    print("- Release")
    windows_zip = base / "Windows" / f"{PROJECT_NAME}_Windows.zip"
    if windows_zip.exists():
        print("\nZIP Windows listo para subir a GitHub Release:")
        print(windows_zip)
    else:
        print("\nPendiente: no se encontró el ejecutable Windows.")
        print("Ejecuta primero: GENERAR_EXE_WINDOWS.bat")
    input("\nPulsa ENTER para volver al menú...")
