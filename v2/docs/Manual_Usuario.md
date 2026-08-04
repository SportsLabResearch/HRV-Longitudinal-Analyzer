# User Manual

## Introduction

HRV-Longitudinal-Analyzer is designed to automate the organization and longitudinal analysis of heart rate variability (HRV) data exported from Kubios HRV Software.

The public version allows researchers, teachers and students to analyse complete datasets while preserving participant confidentiality through automatic anonymization.

---

## Before starting

Verify that the following elements are present in the project folder:

- main_public.py
- config.py
- Datos_publicos
- Resultados_publicos

---

## Preparing the data

Copy all Kubios HRV reports into the **Datos_publicos** folder.

The software automatically detects all available files during execution.

No manual file organization is required.

---

## Running the software

Execute:

```text
main_public.py
```

The program automatically:

- Reads all available HRV reports.
- Detects every participant.
- Organizes repeated measurements.
- Replaces participant identifiers with anonymous labels.
- Processes the complete dataset.
- Generates the corresponding outputs.

---

## Generated results

All outputs are automatically saved inside:

```text
Resultados_publicos
```

Depending on the analysed dataset, the software may generate tables, longitudinal databases and additional scientific outputs.

---

## Participant anonymization

The public version never exports personal identifiers.

Every participant is automatically renamed using anonymous labels:

```text
Subject_001
Subject_002
Subject_003
...
```

This procedure allows datasets to be shared without revealing participant identity.

---

## Typical workflow

1. Export HRV reports from Kubios.
2. Copy the reports into **Datos_publicos**.
3. Run **main_public.py**.
4. Retrieve the generated outputs from **Resultados_publicos**.

---

## Recommendations

- Keep the original Kubios reports unchanged.
- Do not modify automatically generated files.
- Preserve the project folder structure.
- Use the generated outputs for research, teaching or collaborative projects.

