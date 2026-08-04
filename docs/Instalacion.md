# Installation

## Requirements

HRV-Longitudinal-Analyzer requires:

- Python 3.11 or later
- Windows operating system
- Kubios HRV reports exported in the supported format

---

## Installation

1. Download the project.
2. Extract the compressed file.
3. Install the required Python packages:

```text
pip install -r requirements.txt
```

---

## Project structure

The project folder should contain:

```text
main_public.py
config.py
requirements.txt
Datos_publicos/
Resultados_publicos/
docs/
```

---

## Preparing the data

Copy the exported Kubios HRV reports into:

```text
Datos_publicos
```

No additional configuration is required.

---

## Running the software

Execute:

```text
python main_public.py
```

or simply run:

```text
main_public.py
```

depending on your Python installation.

---

## Generated outputs

All generated files are automatically stored inside:

```text
Resultados_publicos
```

The original HRV reports are never modified.

---

## Updating the software

Replace the project files with the latest release while preserving your input and output folders if required.

