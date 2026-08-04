# HRV-Longitudinal-Analyzer

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