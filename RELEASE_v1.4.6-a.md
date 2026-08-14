# HRV-Longitudinal-Analyzer v1.4.6-a

## Cambio principal

Se sustituye el generador de ejecutable anterior por un sistema limpio basado en entorno virtual aislado.

## Corrección

El build anterior podía arrastrar librerías instaladas en el Python global del ordenador, como torch, tensorflow, torchvision u onnxruntime. Esta versión no analiza el entorno global y solo instala las dependencias reales del proyecto.

## Resultado esperado

- Build más rápido.
- Menos dependencias innecesarias.
- Ejecutable más estable.
- Distribución Windows en carpeta.
