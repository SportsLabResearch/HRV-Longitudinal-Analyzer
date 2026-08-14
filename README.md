# HRV-Longitudinal-Analyzer

Current version: **v2.1.0**.

Scientific Python project for organizing, processing and analyzing longitudinal heart rate variability (HRV) data.

## Purpose

HRV-Longitudinal-Analyzer is designed to build clean longitudinal databases from repeated HRV recordings collected across days, weeks or months.

The project is especially useful for studies where HRV is monitored repeatedly in the same participants, such as health, exercise, recovery, training adaptation or clinical follow-up projects.

## Data source

This project can work with HRV files exported from different HRV analysis platforms.

In its first version, the workflow is focused on files exported from Kubios HRV Software. Kubios is used only as a data source. This project is independent and is not affiliated with Kubios.

## Main variables

The core variables considered in the project are:

- HR_bpm
- RMSSD_ms
- LnRMSSD
- RR_medio_ms
- SDNN_ms
- SD1_ms
- SD2_ms
- indice_estres
- frecuencia_respiratoria_resp_min
- LF_ms2
- HF_ms2
- LF_nu_pct
- HF_nu_pct
- LF_HF_ratio

## Private and public data

The local workflow reads private participant folders from `Datos/`. When
`main.py` starts, it automatically creates or updates `Datos_publicos/` using
stable identifiers such as `Sujeto_001`, `Sujeto_002`, and so on. The private
name-to-code correspondence remains only on the local computer and is excluded
from Git.

`Datos/`, `.private/`, `Datos_publicos/` y los resultados generados son
exclusivamente locales y están excluidos del repositorio y de los paquetes de
distribución. La seudonimización del nombre de carpeta no garantiza por sí sola
que una imagen carezca de información identificable.

## Entry points

- `main.py`: complete local workflow for authorized private data.
- `main_public.py`: restricted workflow that reads only `Datos_publicos/` and
  writes only to `Resultados_publicos/`.
- `launcher_hrv.py`: local launcher that verifies Python dependencies before
  starting the complete workflow.

## Analysis date range

When the Kubios OCR analysis starts, the program requests a common initial
date and an optional final date. Both limits are inclusive.

- If a final date is entered, the same period is applied to every selected
  participant.
- If the final date is left blank, each participant is analyzed up to the
  date of her own latest dated image file. Therefore, participants may have
  different final dates.
- Images whose filename does not contain a recognizable date are excluded
  from the period and reported on screen.

The effective initial and final dates are recorded in the report metadata,
the execution summaries and the generated Word document.

## Project structure

```text
HRV-Longitudinal-Analyzer/
│
├── main.py
├── config.py
├── README.md
├── CHANGELOG.md
├── LICENSE
├── requirements.txt
├── .gitignore
│
├── data/
│   ├── raw/
│   ├── processed/
│   └── reports/
│
├── modules/
│   ├── import_kubios.py
│   ├── database.py
│   ├── quality_control.py
│   ├── statistics.py
│   ├── graphics.py
│   ├── reports.py
│   └── utils.py
│
├── docs/
├── examples/
└── tests/
