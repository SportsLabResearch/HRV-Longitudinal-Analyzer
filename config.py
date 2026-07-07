import re

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