# HRV-Longitudinal-Analyzer v1.4.5-a

## Objetivo

Generador de ejecutable Windows optimizado.

## Cambios

- El generador ya no instala `requirements.txt` completo.
- Comprueba únicamente dependencias reales del proyecto.
- Excluye librerías pesadas no usadas: torch, tensorflow, onnxruntime, torchvision, pytest, jupyter, etc.
- Reduce el tiempo de creación del ejecutable.
- Mantiene el launcher local para ejecutar el programa en el ordenador del usuario.

## Uso

Doble clic en:

`GENERAR_EXE_WINDOWS.bat`

El resultado aparecerá en:

`Distribucion/Windows/HRV-Longitudinal-Analyzer.exe`
