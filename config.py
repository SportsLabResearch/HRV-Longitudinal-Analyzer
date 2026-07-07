"""
=========================================================
HRV-Longitudinal-Analyzer
Configuration file
=========================================================
"""

from pathlib import Path

# =====================================================
# PROJECT INFORMATION
# =====================================================

PROJECT_NAME = "HRV-Longitudinal-Analyzer"
VERSION = "0.1.0"
AUTHOR = "José Pino-Ortega"

# =====================================================
# PROJECT PATHS
# =====================================================

ROOT = Path(__file__).resolve().parent

DATA = ROOT / "data"
RAW_DATA = DATA / "raw"
PROCESSED_DATA = DATA / "processed"
REPORTS = DATA / "reports"

DOCS = ROOT / "docs"
EXAMPLES = ROOT / "examples"

# =====================================================
# DATE FORMAT
# =====================================================

DATE_FORMAT = "%Y-%m-%d"

# =====================================================
# HRV VARIABLES
# =====================================================

HRV_VARIABLES = [

    "HR_bpm",

    "RMSSD_ms",

    "LnRMSSD",

    "RR_medio_ms",

    "SDNN_ms",

    "SD1_ms",

    "SD2_ms",

    "indice_estres",

    "frecuencia_respiratoria_resp_min",

    "LF_ms2",

    "HF_ms2",

    "LF_nu_pct",

    "HF_nu_pct",

    "LF_HF_ratio"

]

# =====================================================
# EXCEL OUTPUT
# =====================================================

EXCEL_DATABASE_NAME = "HRV_Longitudinal_Database.xlsx"

# =====================================================
# WORD REPORT
# =====================================================

WORD_REPORT_NAME = "HRV_Report.docx"

# =====================================================
# CONSOLE
# =====================================================

LINE = "=" * 90