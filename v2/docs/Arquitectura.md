# Software Architecture

## Overview

HRV-Longitudinal-Analyzer has been developed as a modular scientific application for organizing, processing and analysing longitudinal heart rate variability (HRV) data.

The public version has been designed to provide a simple and reproducible workflow while preserving participant confidentiality.

---

## Main components

### main_public.py

Main application entry point.

Responsible for controlling the complete workflow, from data import to result generation.

---

### config.py

Contains the project configuration parameters used during execution.

---

### Datos_publicos

Input directory containing the HRV reports exported from Kubios HRV Software.

The software automatically detects all compatible files.

---

### Resultados_publicos

Output directory.

All generated datasets and reports are automatically stored here.

---

## Processing workflow

The software performs the following steps automatically:

1. Read the available HRV reports.
2. Detect all participants.
3. Organize repeated measurements.
4. Apply automatic participant anonymization.
5. Generate standardized outputs.

---

## Anonymization

The public version replaces every participant identifier with anonymous labels:

```text
Subject_001
Subject_002
Subject_003
...
```

No personal identifiers are included in the generated outputs.

---

## Design principles

The software has been developed following the principles of:

- Simplicity
- Reproducibility
- Standardization
- Open Science
- Data confidentiality

