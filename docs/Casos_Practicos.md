# Practical Examples

## Example 1 — Single participant

### Input

Place one or more Kubios HRV reports from the same participant inside:

```text
Datos_publicos
```

### Processing

Run:

```text
python main_public.py
```

The software automatically:

- Imports all reports.
- Detects the participant.
- Organizes repeated measurements.
- Replaces the participant identifier with an anonymous label.

### Output

The generated files are saved inside:

```text
Resultados_publicos
```

---

## Example 2 — Multiple participants

### Input

Copy HRV reports from several participants into:

```text
Datos_publicos
```

The reports may correspond to different assessment days or follow-up sessions.

### Processing

The software automatically:

- Detects every participant.
- Groups repeated measurements.
- Assigns anonymous identifiers (Subject_001, Subject_002, ...).
- Processes the complete longitudinal dataset.

### Output

All processed information is exported to:

```text
Resultados_publicos
```

---

## Example 3 — Scientific collaboration

The public version allows researchers to share datasets without exposing participant identities.

Only anonymized information is included in the generated outputs, facilitating reproducible analyses and collaborative research.

---

## Typical workflow

1. Export HRV reports from Kubios HRV Software.
2. Copy the reports into **Datos_publicos**.
3. Run **main_public.py**.
4. Review the generated outputs in **Resultados_publicos**.

