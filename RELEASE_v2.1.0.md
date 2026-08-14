# HRV-Longitudinal-Analyzer v2.1.0

## Principales mejoras

- Arquitectura modular para participantes, auditoría, validación y distribución.
- Selección de fecha inicial y final para el análisis longitudinal.
- Cálculo e integración de LnRMSSD en los resultados.
- Creación estable de códigos de participante y estructura pública separada.
- Informes organizados por participante y seguimiento global.
- Preparación automática de paquetes de código y compilaciones para Windows.
- Exclusión reforzada de datos privados, resultados y archivos temporales.
- Pruebas de estructura, importación, auditoría y estabilidad de códigos.

## Compatibilidad

Esta versión parte de `v2.0.0` y conserva `main_public.py` como punto de entrada
público. El flujo local completo se ejecuta mediante `main.py` o los lanzadores
para Windows.

## Privacidad

Los paquetes de distribución no contienen imágenes, correspondencias privadas
ni resultados generados. La publicación de conjuntos de datos requiere una
revisión independiente de consentimiento, metadatos y contenido visible.
