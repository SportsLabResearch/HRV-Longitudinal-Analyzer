# CHANGELOG

## v1.4.0-a - Scientific Validation Framework base

### Added
- Nueva estructura `Validation/` para validación técnica y científica.
- Nuevo módulo `modules/validation.py`.
- Nuevo menú principal: `6. Validación científica (SVF)`.
- Dashboard básico de validación científica.
- Estado de reproducibilidad con sistema operativo, Python y commit Git.
- Carpetas base para Technical, Scientific, Benchmark, Reproducibility, Datasets, Publications, Protocols y Results.

### Changed
- Versión del proyecto actualizada a `v1.4.0-a`.
- Próximo objetivo actualizado a `v1.4.0-b - Validación técnica OCR`.


## v1.3.5 - Estable

Fecha: 2026-07-08

### Tipo
- Versión estable de cierre de la serie 1.3.x.

### Cambios
- Consolidación del dashboard del proyecto.
- Consolidación del gestor de participantes.
- Consolidación de la selección activa de participantes.
- Actualización de versión a v1.3.5.
- Preparación para etiquetar esta versión como estable en GitHub.

### Validación
- `main.py` compila correctamente.
- `config.py` compila correctamente.
- `modules/participants.py` compila correctamente.
- No se añaden cambios funcionales grandes para evitar retrocesos.

## v1.3.4
- Conectada la selección activa de participantes con el inicio del análisis OCR.
- Añadido resumen preanálisis de participantes activos.
- Actualizada versión del proyecto a v1.3.4.

## v1.3.3 - Selección activa de participantes
- Añadida selección activa persistente de participantes.
- El gestor guarda la selección en Resultados/Participantes/seleccion_activa_participantes.json.
- El análisis Kubios OCR puede reutilizar la selección activa antes de pedir una nueva selección.
- Añadida opción de menú para ver la selección activa.
- Añadidas opciones en el gestor para ver o borrar la selección activa.

## v1.3.2
- Mejora del gestor de participantes.
- Añadido detalle individual por participante.
- Añadida detección básica de fechas desde nombres de archivo.
- Añadido cálculo de sesiones posibles por fecha.
- Añadida exportación del listado de participantes a Excel en Resultados/Participantes/.
- Añadida opción de volver al menú pulsando ENTER.

## v1.3.1-fix
- Corrige la carga del gestor de participantes desde el menú.
- Añade importación robusta del módulo `modules/participants.py`.
- Añade mensaje técnico si el módulo no se encuentra.


## v1.3.1
- Añadido nuevo módulo `modules/participants.py`.
- Añadido gestor interactivo de participantes desde el menú principal.
- Selección compatible con un participante, varios, intervalo o todos.
- Detección automática de participantes desde `Datos/`.
- Conteo básico de archivos, imágenes y tablas por participante.
- Actualizada versión visible a v1.3.1.

## v1.3.0-d
- Añadida sección de estado técnico al dashboard.
- Añadida comprobación resumida de Git.
- Añadido resumen de requirements.txt.
- Mejorada la sección de calidad del proyecto.
- Actualizada versión a v1.3.0-d.

## v1.3.0-c
- Dashboard inteligente con estadísticas reales del repositorio.
- Conteo automático de archivos Python, líneas útiles, funciones y clases.
- Conteo automático de sesiones/archivos de datos, informes Word, Excel, PDF y gráficos.
- Nueva sección de calidad del proyecto: Git, README, LICENSE, CHANGELOG, requirements, tests y módulo dashboard.
- Progreso global calculado dinámicamente a partir del estado del proyecto y calidad básica.


## v1.3.0-b
- Corrección del dashboard del proyecto.
- Actualización de versión visible a v1.3.0-b.
- Corrección de la nota final para evitar salida como tupla.
- Actualización del próximo objetivo a v1.3.0-c.


## v1.3.0
- Dashboard profesional (planificado)
- Preparación para modularización.
- Hoja de ruta de arquitectura.

