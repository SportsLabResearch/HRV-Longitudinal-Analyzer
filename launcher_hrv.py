# -*- coding: utf-8 -*-
"""
HRV-Longitudinal-Analyzer Launcher Local v2.1.0
Comprueba el ordenador donde se ejecuta, instala solo dependencias Python ausentes
según requirements.txt y lanza main.py.
"""
from __future__ import annotations

import importlib.util
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from datetime import datetime

ROOT = Path(__file__).resolve().parent
REQ_FILE = ROOT / "requirements.txt"
MAIN_FILE = ROOT / "main.py"
LOG_DIR = ROOT / "Resultados_Auditoria" / "Launcher"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / f"launcher_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"

PACKAGE_IMPORTS = {
    "python-docx": "docx",
    "opencv-python": "cv2",
    "pyyaml": "yaml",
    "pillow": "PIL",
}


def log(msg: str) -> None:
    print(msg)
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(msg + "\n")


def clean_requirement(line: str) -> str | None:
    line = line.strip()
    if not line or line.startswith("#"):
        return None
    line = line.split(";")[0].strip()
    name = re.split(r"[<>=!~\[]", line, maxsplit=1)[0].strip()
    return name or None


def import_name(package: str) -> str:
    return PACKAGE_IMPORTS.get(package.lower(), package.replace("-", "_"))


def is_installed(package: str) -> bool:
    return importlib.util.find_spec(import_name(package)) is not None


def read_requirements() -> list[str]:
    if not REQ_FILE.exists():
        return []
    packages = []
    for line in REQ_FILE.read_text(encoding="utf-8", errors="ignore").splitlines():
        pkg = clean_requirement(line)
        if pkg and pkg.lower() != "pyinstaller":
            packages.append(pkg)
    seen = []
    for p in packages:
        if p.lower() not in [x.lower() for x in seen]:
            seen.append(p)
    return seen


def install_missing(packages: list[str]) -> bool:
    if not packages:
        return True
    log("\nInstalando solo dependencias ausentes:")
    for p in packages:
        log(f" - {p}")
    cmd = [sys.executable, "-m", "pip", "install", *packages]
    result = subprocess.run(cmd, cwd=str(ROOT))
    return result.returncode == 0


def check_tesseract() -> None:
    exe = shutil.which("tesseract")
    if exe:
        log(f"OK         Tesseract OCR detectado: {exe}")
    else:
        log("AVISO      Tesseract OCR no detectado en PATH")
        log("           Si el OCR falla, instala Tesseract y añádelo al PATH.")


def launch_main() -> int:
    if not MAIN_FILE.exists():
        log("ERROR      No se encuentra main.py")
        return 1
    log("\nEjecutando HRV-Longitudinal-Analyzer...")
    return subprocess.call([sys.executable, str(MAIN_FILE)], cwd=str(ROOT))


def main() -> int:
    log("=" * 62)
    log(" HRV-LONGITUDINAL-ANALYZER | LAUNCHER LOCAL v2.1.0")
    log("=" * 62)
    log(f"Carpeta : {ROOT}")
    log(f"Python  : {sys.version.split()[0]}")
    log(f"Ejecutor: {sys.executable}")
    log("-" * 62)

    packages = read_requirements()
    if not packages:
        log("AVISO      No se ha encontrado requirements.txt o está vacío")
    else:
        missing = []
        for p in packages:
            status = "OK" if is_installed(p) else "FALTA"
            log(f"{status:<10} {p}")
            if status == "FALTA":
                missing.append(p)
        if missing:
            ok = install_missing(missing)
            if not ok:
                log("ERROR      No se pudieron instalar todas las dependencias")
                input("Pulsa ENTER para salir...")
                return 1
        else:
            log("\nNo falta ninguna dependencia Python.")

    check_tesseract()
    log(f"\nLog: {LOG_FILE}")
    log("=" * 62)
    return launch_main()


if __name__ == "__main__":
    raise SystemExit(main())
