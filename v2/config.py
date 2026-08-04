import re
from pathlib import Path

# ======================================================
# CONFIGURACIÓN GLOBAL DEL PROYECTO
# ======================================================


# ======================================================
# RUTAS Y CARPETAS ESTÁNDAR
# ======================================================

SCRIPT_DIR = Path(__file__).resolve().parent
DATOS_DIR_NAME = "Datos"
RESULTADOS_DIR_NAME = "Resultados"
ASSETS_DIR_NAME = "assets"
SHARED_IMAGES_SUBDIR = "images"

RESULTADOS_DIR = SCRIPT_DIR / RESULTADOS_DIR_NAME
ASSETS_DIR = SCRIPT_DIR / ASSETS_DIR_NAME
SHARED_IMAGES_DIR = ASSETS_DIR / SHARED_IMAGES_SUBDIR

PARTICIPANTES_DIR_NAMES = (
    "Datos",
    "DATOS",
    "datos",
    "Dato",
    "DATO",
    "dato",
    "analisis de participantes",
    "análisis de participantes",
    "analisis_participantes",
    "participantes",
)

IMAGE_EXTENSIONS = ("*.png", "*.jpg", "*.jpeg", "*.bmp", "*.webp")

DATE_PATTERNS = [
    re.compile(r"(?P<date>\d{4}-\d{2}-\d{2})\s+at\s+(?P<time>\d{2}\.\d{2}\.\d{2})", re.IGNORECASE),
    re.compile(r"(?P<date>\d{4}-\d{2}-\d{2})[ _-]+(?P<time>\d{2}[\.:\-]\d{2}[\.:\-]\d{2})", re.IGNORECASE),
    re.compile(r"(?P<date>\d{2}[\-\.]\d{2}[\-\.]\d{4})[ _-]+(?P<time>\d{2}[\.:\-]\d{2}[\.:\-]\d{2})", re.IGNORECASE),
    re.compile(r"(?P<date>\d{4}-\d{2}-\d{2})", re.IGNORECASE),
    re.compile(r"(?P<date>\d{2}[\-\.]\d{2}[\-\.]\d{4})", re.IGNORECASE),
]

SESSION_GROUP_GAP_SECONDS = 90


# ======================================================
# IDENTIDAD Y ESTADO DEL PROYECTO
# ======================================================

PROJECT_VERSION = "v1.4.2-a"
PROJECT_NEXT_OBJECTIVE = "v1.4.2-estable - Publicación final en main"

# ======================================================
# ESTADO DEL PROYECTO
# ======================================================

PROJECT_PROGRESS = {
    "Funcionalidad original V41 integrada": 100,
    "Git y estructura base": 100,
    "Configuración básica": 100,
    "Gestión de participantes": 100,
    "Estructura de resultados": 100,
    "Compatibilidad multiplataforma": 100,
    "Modularización": 100,
    "Documentación": 100,
    "Pruebas": 100,
    "Versión preparada para estable": 100,
}

PROJECT_OVERALL_PROGRESS = 100
PROJECT_PROGRESS_NOTES = """La versión actual está completa para revisión final.
La publicación estable se realizará al fusionar dev en main y crear el release correspondiente."""
