# -*- coding: utf-8 -*-
"""
Scientific Validation Framework (SVF)
HRV-Longitudinal-Analyzer v1.4.0-a

Este módulo crea y muestra el estado básico del sistema de validación
científica del proyecto. Está diseñado para no depender de paquetes externos
ni alterar el flujo principal del análisis OCR.
"""

from __future__ import annotations

import platform
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, Tuple

VALIDATION_DIR_NAME = "Validation"

VALIDATION_STRUCTURE = (
    "01_Technical/OCR",
    "01_Technical/Dashboard",
    "01_Technical/Participants",
    "01_Technical/Reports",
    "01_Technical/Graphics",
    "01_Technical/Export",
    "01_Technical/Performance",
    "01_Technical/Compatibility",
    "02_Scientific/OCR",
    "02_Scientific/HRV",
    "02_Scientific/Longitudinal",
    "02_Scientific/Statistics",
    "02_Scientific/Reports",
    "02_Scientific/Algorithms",
    "03_Benchmark",
    "04_Reproducibility",
    "05_Datasets/OCR",
    "05_Datasets/HRV",
    "05_Datasets/Longitudinal",
    "06_Publications/OCR_Validation/figures",
    "06_Publications/OCR_Validation/supplementary",
    "07_Protocols",
    "Results",
)

VALIDATION_AREAS = {
    "Validación técnica": 20,
    "Validación científica": 5,
    "Benchmark": 5,
    "Reproducibilidad": 15,
    "Datasets de validación": 10,
    "Protocolos": 10,
    "Publicaciones": 0,
}


def _bar(percent: int, width: int = 20) -> str:
    percent = max(0, min(100, int(percent)))
    filled = round(width * percent / 100)
    return "█" * filled + "░" * (width - filled)


def _ok(value: bool) -> str:
    return "OK" if value else "NO ENCONTRADO"


def _project_root(project_root: Path | str | None = None) -> Path:
    if project_root is None:
        return Path(__file__).resolve().parents[1]
    return Path(project_root).resolve()


def get_validation_root(project_root: Path | str | None = None) -> Path:
    return _project_root(project_root) / VALIDATION_DIR_NAME


def ensure_validation_structure(project_root: Path | str | None = None) -> Path:
    """Crea la estructura mínima de validación si no existe."""
    validation_root = get_validation_root(project_root)
    validation_root.mkdir(parents=True, exist_ok=True)

    for relative in VALIDATION_STRUCTURE:
        folder = validation_root / relative
        folder.mkdir(parents=True, exist_ok=True)
        gitkeep = folder / ".gitkeep"
        if not gitkeep.exists():
            gitkeep.write_text("", encoding="utf-8")

    readme = validation_root / "README.md"
    if not readme.exists():
        readme.write_text(
            "# Scientific Validation Framework (SVF)\n\n"
            "Estructura base para la validación técnica y científica del software.\n",
            encoding="utf-8",
        )

    return validation_root


def _count_files(path: Path, patterns: Iterable[str]) -> int:
    total = 0
    if not path.exists():
        return 0
    for pattern in patterns:
        total += sum(1 for p in path.rglob(pattern) if p.is_file())
    return total


def _git_commit(project_root: Path) -> str:
    try:
        result = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=str(project_root),
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass
    return "No disponible"


def validation_status(project_root: Path | str | None = None) -> Dict[str, object]:
    root = _project_root(project_root)
    validation_root = ensure_validation_structure(root)

    required = [validation_root / relative for relative in VALIDATION_STRUCTURE]
    existing = sum(1 for p in required if p.exists())
    structure_percent = round(existing * 100 / len(required)) if required else 0

    results_root = validation_root / "Results"
    datasets_root = validation_root / "05_Datasets"
    protocols_root = validation_root / "07_Protocols"
    publications_root = validation_root / "06_Publications"

    return {
        "project_root": root,
        "validation_root": validation_root,
        "required_folders": len(required),
        "existing_folders": existing,
        "structure_percent": structure_percent,
        "word_reports": _count_files(validation_root, ("*.docx",)),
        "excel_reports": _count_files(validation_root, ("*.xlsx", "*.xls")),
        "figures": _count_files(validation_root, ("*.png", "*.jpg", "*.jpeg")),
        "datasets": _count_files(datasets_root, ("*.xlsx", "*.xls", "*.csv", "*.json")),
        "protocols": _count_files(protocols_root, ("*.md", "*.docx", "*.pdf")),
        "publication_files": _count_files(publications_root, ("*.md", "*.docx", "*.xlsx", "*.bib")),
        "result_files": _count_files(results_root, ("*.*",)),
        "python": sys.version.split()[0],
        "system": f"{platform.system()} {platform.release()}",
        "git_commit": _git_commit(root),
    }


def print_validation_dashboard(project_root: Path | str | None = None) -> None:
    status = validation_status(project_root)

    print("\n" + "=" * 70)
    print(" SCIENTIFIC VALIDATION FRAMEWORK - DASHBOARD")
    print("=" * 70)

    print("\nPROYECTO")
    print("-" * 70)
    print(f"Directorio del proyecto:          {status['project_root']}")
    print(f"Directorio Validation:            {status['validation_root']}")
    print(f"Sistema operativo:                {status['system']}")
    print(f"Python:                           {status['python']}")
    print(f"Git commit:                       {status['git_commit']}")

    print("\nESTRUCTURA")
    print("-" * 70)
    print(f"Carpetas requeridas:              {status['required_folders']}")
    print(f"Carpetas existentes:              {status['existing_folders']}")
    print(f"Estructura SVF:                   {status['structure_percent']}%  {_bar(status['structure_percent'])}")

    print("\nEVIDENCIAS DE VALIDACIÓN")
    print("-" * 70)
    print(f"Informes Word:                    {status['word_reports']}")
    print(f"Archivos Excel:                   {status['excel_reports']}")
    print(f"Figuras:                          {status['figures']}")
    print(f"Datasets de validación:           {status['datasets']}")
    print(f"Protocolos:                       {status['protocols']}")
    print(f"Materiales de publicación:        {status['publication_files']}")
    print(f"Resultados históricos:            {status['result_files']}")

    print("\nÁREAS DE VALIDACIÓN")
    print("-" * 70)
    for area, percent in VALIDATION_AREAS.items():
        print(f"{area:<35} {percent:>3}%  {_bar(percent)}")

    print("\nCARPETAS PRINCIPALES")
    print("-" * 70)
    for folder in (
        "01_Technical",
        "02_Scientific",
        "03_Benchmark",
        "04_Reproducibility",
        "05_Datasets",
        "06_Publications",
        "07_Protocols",
        "Results",
    ):
        path = status["validation_root"] / folder
        print(f"{folder + '/':<35} {_ok(path.exists())}")

    print("\nNota:")
    print("SVF v1.4.0-a crea la estructura base de validación técnica y científica.")
    print("Las siguientes versiones añadirán validación OCR, reproducibilidad y benchmark.")


def show_validation_menu(project_root: Path | str | None = None) -> None:
    root = _project_root(project_root)
    ensure_validation_structure(root)

    while True:
        print("\n" + "=" * 70)
        print(" VALIDACIÓN CIENTÍFICA - SVF")
        print("=" * 70)
        print("1. Estado general de validación")
        print("2. Crear/ver estructura Validation")
        print("3. Estado de reproducibilidad")
        print("0. Volver")

        option = input("\nSelecciona una opción: ").strip()

        if option == "1":
            print_validation_dashboard(root)
            input("\nPulsa ENTER para volver al menú de validación...")
            continue

        if option == "2":
            validation_root = ensure_validation_structure(root)
            print("\nEstructura Validation preparada correctamente.")
            print(f"Ruta: {validation_root}")
            input("\nPulsa ENTER para volver al menú de validación...")
            continue

        if option == "3":
            status = validation_status(root)
            print("\n" + "=" * 70)
            print(" REPRODUCIBILIDAD")
            print("=" * 70)
            print(f"Fecha:                 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Sistema:               {status['system']}")
            print(f"Python:                {status['python']}")
            print(f"Git commit:            {status['git_commit']}")
            print(f"Proyecto:              {status['project_root']}")
            input("\nPulsa ENTER para volver al menú de validación...")
            continue

        if option == "0":
            break

        print("Opción no válida.")


if __name__ == "__main__":
    show_validation_menu()
