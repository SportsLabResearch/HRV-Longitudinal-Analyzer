# RELEASE v1.4.1-a

## Tipo de versión
Versión alfa funcional de cierre y control interno del proyecto.

## Cambios principales
- Integrada la opción `7. Auditoría del proyecto` en el menú principal.
- Añadido el módulo `modules/auditoria.py`.
- Creación automática de la carpeta `Resultados_Auditoria/`.
- Generación de informe de auditoría en Word `.docx`.
- Generación de informe de auditoría en Excel `.xlsx`.
- Generación de resumen `.txt`.
- Comprobación automática de archivos, carpetas, módulos, documentación, releases y tests.
- No modifica el análisis OCR ni la gestión de participantes.

## Estado
Funcional y probado por sintaxis.

## Nota
La auditoría indica si el proyecto está preparado o no para convertirse en versión estable.

## Actualización de cierre

- Añadida carpeta `tests/` con pruebas automáticas básicas reales.
- Añadidas comprobaciones de estructura del proyecto.
- Añadidas comprobaciones de compilación de `main.py`, `config.py` y módulos.
- Añadida prueba de auditoría y generación de resultados.
- La auditoría del proyecto ya puede detectar tests Python reales.
