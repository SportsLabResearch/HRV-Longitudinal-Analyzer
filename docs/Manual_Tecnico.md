# Technical Manual

## Overview

HRV-Longitudinal-Analyzer has been developed to automate the organization and processing of longitudinal heart rate variability (HRV) datasets exported from Kubios HRV Software.

The public version performs all processing automatically while preserving participant confidentiality.

---

## Input data

The software reads all compatible HRV reports located inside:

```text
Datos_publicos
```

No manual configuration of the files is required.

---

## Processing pipeline

During execution the software automatically performs the following operations:

1. File detection.
2. Data import.
3. Participant identification.
4. Longitudinal organization of repeated measurements.
5. Automatic participant anonymization.
6. Data validation.
7. Output generation.

---

## Participant anonymization

Each detected participant is assigned a unique anonymous identifier.

Example:

```text
Subject_001
Subject_002
Subject_003
```

The original participant names are never included in the generated outputs.

---

## Outputs

All generated files are stored inside:

```text
Resultados_publicos
```

The original HRV reports remain unchanged.

---

## Software philosophy

The public version has been designed to provide:

- Standardized processing.
- Automatic anonymization.
- Reproducible analyses.
- Scientific transparency.
- Easy distribution for research and teaching.

