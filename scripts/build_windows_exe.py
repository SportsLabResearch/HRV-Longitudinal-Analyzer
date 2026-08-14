# -*- coding: utf-8 -*-
"""
Generador Windows optimizado v1.4.5-a.
No instala requirements.txt completo.
No usa --collect-all genérico.
Excluye librerías pesadas no usadas por HRV-Longitudinal-Analyzer.
"""
from __future__ import annotations

import importlib.util
import platform
import shutil
import subprocess
import sys
from pathlib import Path
from datetime import datetime

PROJECT_NAME = "HRV-Longitudinal-Analyzer"
VERSION = "v1.4.5-a"
ROOT = Path(__file__).resolve().parents[1]
DIST_FINAL = ROOT / "Distribucion" / "Windows"
BUILD_DIR = ROOT / "build"
DIST_DIR = ROOT / "dist"
EXE_NAME = f"{PROJECT_NAME}.exe"
LOG_DIR = ROOT / "Resultados_Auditoria" / "Build_EXE"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / f"build_exe_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"

# Dependencias reales usadas por el proyecto.
# No se instalan librerías globales ni paquetes innecesarios.
REQUIRED_PACKAGES = {
    "numpy": "numpy",
    "pandas": "pandas",
    "matplotlib": "matplotlib",
    "openpyxl": "openpyxl",
    "python-docx": "docx",
    "opencv-python": "cv2",
    "pytesseract": "pytesseract",
    "pyinstaller": "PyInstaller",
}

# Paquetes pesados que NO pertenecen al proyecto y que PyInstaller puede detectar
# por estar instalados en el Python del ordenador.
EXCLUDES = [
    "torch", "torchvision", "torchaudio", "tensorflow", "keras", "onnx", "onnxruntime",
    "tensorboard", "transformers", "sklearn", "scikit-learn", "pytest",
    "IPython", "jupyter", "notebook", "pygame", "pyarrow", "fsspec",
    "numba", "llvmlite", "dask", "ray", "xgboost", "lightgbm",
]

HIDDEN_IMPORTS = [
    "docx", "openpyxl", "pytesseract", "cv2",
    "matplotlib.backends.backend_agg",
]

COLLECT_SUBMODULES = [
    "docx",
    "openpyxl",
]

COLLECT_DATA = [
    "matplotlib",
]


def log(msg: str) -> None:
    print(msg)
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(msg + "\n")


def run(cmd: list[str]) -> None:
    log(" ".join(str(c) for c in cmd))
    subprocess.check_call(cmd, cwd=str(ROOT))


def module_installed(import_name: str) -> bool:
    return importlib.util.find_spec(import_name) is not None


def install_missing() -> None:
    missing = []
    log("\nComprobando dependencias reales del proyecto:")
    for package, import_name in REQUIRED_PACKAGES.items():
        ok = module_installed(import_name)
        log(f"{'OK' if ok else 'FALTA':<8} {package}")
        if not ok:
            missing.append(package)
    if missing:
        log("\nInstalando únicamente dependencias ausentes:")
        for m in missing:
            log(f" - {m}")
        run([sys.executable, "-m", "pip", "install", *missing])
    else:
        log("\nNo falta ninguna dependencia Python.")


def check_project() -> None:
    for name in ["main.py", "config.py"]:
        if not (ROOT / name).exists():
            raise FileNotFoundError(f"No se encuentra {name}")
    run([sys.executable, "-m", "py_compile", "main.py"])
    run([sys.executable, "-m", "py_compile", "config.py"])


def clean_previous() -> None:
    for p in [BUILD_DIR, DIST_DIR]:
        if p.exists():
            shutil.rmtree(p)


def build_exe() -> None:
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--clean",
        "--noconfirm",
        "--name", PROJECT_NAME,
    ]
    for item in EXCLUDES:
        cmd += ["--exclude-module", item]
    for item in HIDDEN_IMPORTS:
        cmd += ["--hidden-import", item]
    for item in COLLECT_SUBMODULES:
        cmd += ["--collect-submodules", item]
    for item in COLLECT_DATA:
        cmd += ["--collect-data", item]
    cmd += ["main.py"]
    run(cmd)


def copy_distribution() -> Path:
    DIST_FINAL.mkdir(parents=True, exist_ok=True)
    exe_src = DIST_DIR / EXE_NAME
    exe_dst = DIST_FINAL / EXE_NAME
    if not exe_src.exists():
        raise FileNotFoundError(f"No se encontró el ejecutable: {exe_src}")
    shutil.copy2(exe_src, exe_dst)

    for name in ["README.md", "LICENSE", "CHANGELOG.md", "RELEASE_v1.4.5-a.md", "README_LAUNCHER_LOCAL.md"]:
        src = ROOT / name
        if src.exists():
            shutil.copy2(src, DIST_FINAL / name)

    for folder in ["docs", "Validation", "tests"]:
        src = ROOT / folder
        dst = DIST_FINAL / folder
        if src.exists():
            if dst.exists():
                shutil.rmtree(dst)
            shutil.copytree(src, dst)
    return exe_dst


def main() -> int:
    log("=" * 70)
    log(f"GENERADOR EJECUTABLE WINDOWS OPTIMIZADO | {VERSION}")
    log("=" * 70)
    log(f"Carpeta : {ROOT}")
    log(f"Python  : {sys.version.split()[0]}")
    log(f"Log     : {LOG_FILE}")
    if platform.system().lower() != "windows":
        log("Este generador debe ejecutarse en Windows para crear el .exe de Windows.")
        return 1

    install_missing()
    check_project()
    clean_previous()
    build_exe()
    exe = copy_distribution()

    log("=" * 70)
    log("EJECUTABLE GENERADO")
    log(str(exe))
    log("=" * 70)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
