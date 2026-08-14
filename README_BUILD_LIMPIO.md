# Build limpio Windows v1.4.6-a

Ejecutar:

`GENERAR_EXE_WINDOWS.bat`

Este generador crea un entorno virtual temporal `.build_env_hrv` e instala solo las dependencias reales del proyecto. No usa `requirements.txt` y no analiza la instalación global de Python.

Excluye librerías pesadas que no pertenecen al proyecto: torch, tensorflow, torchvision, onnxruntime, pyarrow, etc.

El resultado se genera en:

`Distribucion/Windows/HRV-Longitudinal-Analyzer/HRV-Longitudinal-Analyzer.exe`

Es una distribución tipo carpeta: el usuario final ejecuta el `.exe` y no necesita instalar Python.
