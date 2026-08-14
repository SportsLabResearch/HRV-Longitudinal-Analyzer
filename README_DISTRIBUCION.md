# Preparación de distribución oficial

Desde la versión v1.5.0-a el programa incluye la opción:

```text
8. Preparar distribución oficial
```

Esta opción crea automáticamente una carpeta `Distribucion_Release/` con:

- `Codigo_Fuente/`
- `Windows/`
- `Release/`

Si ya existe el ejecutable, genera también `HRV-Longitudinal-Analyzer_Windows.zip`, listo para adjuntar a GitHub Releases.

GitHub genera automáticamente `Source code (zip)` y `Source code (tar.gz)`, por lo que normalmente solo hay que subir el ZIP de Windows.
