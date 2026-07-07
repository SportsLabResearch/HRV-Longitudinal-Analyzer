"""
=========================================================
HRV-Longitudinal-Analyzer
Common longitudinal HRV database model
=========================================================
"""

import pandas as pd

from config import HRV_VARIABLES


# =====================================================
# REQUIRED METADATA COLUMNS
# =====================================================

PARTICIPANT_COLUMNS = [
    "participant_id",
    "participant_name",
    "sex",
    "date_birth",
    "group",
    "center",
]

RECORD_COLUMNS = [
    "record_id",
    "date",
    "time",
    "datetime",
    "device",
    "software",
    "source",
    "position",
    "duration_s",
    "protocol",
    "session",
    "notes",
]

# =====================================================
# OFFICIAL DATABASE SCHEMA
# =====================================================

DATABASE_COLUMNS = PARTICIPANT_COLUMNS + RECORD_COLUMNS + HRV_VARIABLES


def create_empty_database() -> pd.DataFrame:
    """
    Create an empty longitudinal HRV database
    using the official project schema.
    """
    return pd.DataFrame(columns=DATABASE_COLUMNS)


def validate_database_columns(df: pd.DataFrame) -> tuple[bool, list[str]]:
    """
    Validate whether a DataFrame contains all required columns.

    Returns
    -------
    valid : bool
        True if all columns are present.
    missing_columns : list
        List of missing columns.
    """
    missing_columns = [col for col in DATABASE_COLUMNS if col not in df.columns]
    valid = len(missing_columns) == 0

    return valid, missing_columns


def add_record(database: pd.DataFrame, record: dict) -> pd.DataFrame:
    """
    Add one HRV record to the longitudinal database.

    Missing fields are filled with None.
    Extra fields are ignored.
    """
    clean_record = {}

    for column in DATABASE_COLUMNS:
        clean_record[column] = record.get(column, None)

    new_row = pd.DataFrame([clean_record])

    return pd.concat([database, new_row], ignore_index=True)


def export_database_to_excel(df: pd.DataFrame, output_path) -> None:
    """
    Export the longitudinal HRV database to Excel.
    """
    df.to_excel(output_path, index=False)


def database_summary(df: pd.DataFrame) -> dict:
    """
    Generate a basic summary of the longitudinal database.
    """
    summary = {
        "rows": len(df),
        "columns": len(df.columns),
        "participants": df["participant_id"].nunique() if "participant_id" in df.columns else 0,
        "records": df["record_id"].nunique() if "record_id" in df.columns else 0,
        "variables": len(HRV_VARIABLES),
    }

    return summary