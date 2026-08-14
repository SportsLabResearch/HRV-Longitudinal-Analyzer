# -*- coding: utf-8 -*-
"""
Generador limpio de ejecutable Windows v2.1.0.
Crea un entorno virtual aislado y empaqueta solo las dependencias del proyecto.
No usa requirements.txt.
No analiza la instalación global de Python.
No arrastra torch/tensorflow/onnxruntime/torchvision.
"""
from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
import venv
import zipfile
from datetime import datetime
from pathlib import Path

PROJECT_NAME = "HRV-Longitudinal-Analyzer"
VERSION = "v2.1.0"
ROOT = Path(__file__).resolve().parents[1]
BUILD_ENV = ROOT / ".build_env_hrv"
BUILD_DIR = ROOT / "build"
DIST_DIR = ROOT / "dist"
DIST_FINAL = ROOT / "Distribucion" / "Windows"
LOG_DIR = ROOT / "Resultados_Auditoria" / "Build_EXE"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / f"build_limpio_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"

# Solo dependencias reales del proyecto.
REQUIRED = [
    "pyinstaller>=6.0",
    "numpy>=2.0",
    "pandas>=2.3",
    "matplotlib>=3.10",
    "openpyxl>=3.1",
    "xlsxwriter>=3.2",
    "python-docx>=1.2",
    "opencv-python>=4.10",
    "pytesseract>=0.3",
]

EXCLUDES = [
    "torch", "torchvision", "torchaudio", "tensorflow", "keras",
    "onnx", "onnxruntime", "tensorboard", "transformers",
    "sklearn", "scikit-learn", "pytest", "IPython", "jupyter",
    "notebook", "pygame", "pyarrow", "fsspec", "numba",
    "llvmlite", "dask", "ray", "xgboost", "lightgbm",
]

HIDDEN_IMPORTS = [
    "docx", "openpyxl", "pytesseract", "cv2",
    "matplotlib.backends.backend_agg",
]


def log(msg: str) -> None:
    print(msg)
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(msg + "\n")


def run(cmd: list[str], cwd: Path | None = None) -> None:
    log(" ".join(str(x) for x in cmd))
    subprocess.check_call(cmd, cwd=str(cwd or ROOT))


def python_in_venv() -> Path:
    if platform.system().lower() == "windows":
        return BUILD_ENV / "Scripts" / "python.exe"
    return BUILD_ENV / "bin" / "python"


def create_clean_env() -> Path:
    if BUILD_ENV.exists():
        shutil.rmtree(BUILD_ENV)
    log("Creando entorno limpio aislado...")
    venv.EnvBuilder(with_pip=True, clear=True).create(BUILD_ENV)
    py = python_in_venv()
    run([str(py), "-m", "pip", "install", "--upgrade", "pip"])
    log("Instalando solo dependencias reales del proyecto...")
    run([str(py), "-m", "pip", "install", *REQUIRED])
    return py


def check_project(py: Path) -> None:
    for name in ["main.py", "config.py"]:
        if not (ROOT / name).exists():
            raise FileNotFoundError(f"No se encuentra {name}")
    run([str(py), "-m", "py_compile", "main.py"])
    run([str(py), "-m", "py_compile", "config.py"])


def clean_build_dirs() -> None:
    for p in [BUILD_DIR, DIST_DIR]:
        if p.exists():
            shutil.rmtree(p)


def build(py: Path) -> None:
    # ONEDIR es mucho más rápido y fiable que ONEFILE.
    # El usuario final ejecuta el .exe de la carpeta, sin instalar Python.
    cmd = [
        str(py), "-m", "PyInstaller",
        "--onedir",
        "--clean",
        "--noconfirm",
        "--name", PROJECT_NAME,
        "--noupx",
    ]
    for e in EXCLUDES:
        cmd += ["--exclude-module", e]
    for h in HIDDEN_IMPORTS:
        cmd += ["--hidden-import", h]
    cmd += ["main.py"]
    run(cmd)


def copy_final() -> Path:
    DIST_FINAL.mkdir(parents=True, exist_ok=True)
    src_folder = DIST_DIR / PROJECT_NAME
    dst_folder = DIST_FINAL / PROJECT_NAME
    if not src_folder.exists():
        raise FileNotFoundError(f"No se encontró la carpeta generada: {src_folder}")
    if dst_folder.exists():
        shutil.rmtree(dst_folder)
    shutil.copytree(src_folder, dst_folder)

    for name in ["README.md", "LICENSE", "CHANGELOG.md", "RELEASE_v2.1.0.md", "RELEASE_v1.5.0-a.md", "README_BUILD_LIMPIO.md"]:
        src = ROOT / name
        if src.exists():
            shutil.copy2(src, DIST_FINAL / name)
    
    zip_path = DIST_FINAL / f"{PROJECT_NAME}_Windows.zip"
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for p in dst_folder.rglob("*"):
            if p.is_file():
                zf.write(p, dst_folder.name / p.relative_to(dst_folder))
        for extra in ["README.md", "LICENSE", "CHANGELOG.md"]:
            extra_path = DIST_FINAL / extra
            if extra_path.exists():
                zf.write(extra_path, extra_path.name)
    return dst_folder / f"{PROJECT_NAME}.exe"


def main() -> int:
    log("=" * 70)
    log(f"GENERADOR LIMPIO WINDOWS | {VERSION}")
    log("=" * 70)
    log(f"Proyecto : {ROOT}")
    log(f"Log      : {LOG_FILE}")

    if platform.system().lower() != "windows":
        log("Este generador debe ejecutarse en Windows.")
        return 1

    py = create_clean_env()
    check_project(py)
    clean_build_dirs()
    build(py)
    exe = copy_final()

    log("=" * 70)
    log("EJECUTABLE GENERADO")
    log(str(exe))
    log("ZIP WINDOWS GENERADO")
    log(str(DIST_FINAL / f"{PROJECT_NAME}_Windows.zip"))
    log("=" * 70)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
