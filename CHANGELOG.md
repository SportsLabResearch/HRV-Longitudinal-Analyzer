# Mejora de selección del periodo de análisis

- El análisis solicita una fecha inicial y una fecha final opcional.
- El intervalo es inclusivo.
- Si la fecha final se deja vacía, se usa para cada participante la fecha de
  su último archivo de imagen reconocido.
- La fecha final efectiva queda registrada en los metadatos, resúmenes e
  informes Word.
- La codificación pública conserva el mismo `Sujeto_XXX` cuando cambia el
  texto de una carpeta privada que mantiene su número inicial.
- Antes de publicar, se comprueba que `Datos/` y `.private/` estén excluidos
  de Git; la operación prepara únicamente `Datos_publicos/`.

# v1.4.5-a

- Añadido generador Windows optimizado.
- Eliminado uso de `pip install -r requirements.txt` en el generador EXE.
- Añadidas exclusiones PyInstaller para evitar torch, tensorflow, onnxruntime, torchvision y paquetes no usados.
- Añadida documentación `README_BUILD_OPTIMIZADO.md`.


## v1.4.4-a - Launcher local inteligente

### Añadido
- `launcher_hrv.py` para comprobar el ordenador donde se ejecuta.
- `INICIAR_HRV_WINDOWS.bat` para inicio directo en Windows.
- Instalación solo de dependencias ausentes.
- Logs del launcher en `Resultados_Auditoria/Launcher/`.

### Objetivo
- Evitar ejecutables pesados y permitir una distribución práctica para usuarios con Python instalado.

## v1.4.3-a

- Preparada fase de distribución Windows.
- Añadido generador de ejecutable `GENERAR_EXE_WINDOWS.bat`.
- Añadido script `scripts/build_windows_exe.py`.
- Añadida carpeta `Distribucion/Windows/`.
- Versión actualizada a v1.4.3-a.

## v1.4.2-a

### Changed
- Clarificada la función de las opciones 2, 3 y 7 del menú principal.
- La opción 2 queda como Estado del desarrollo.
- La opción 3 queda como Dashboard operativo, sin porcentajes de desarrollo.
- La opción 7 queda como Auditoría del proyecto para revisión estable.
- Progreso de la versión actualizado a 100% tras superar auditoría y tests básicos.

# Changelog

## v1.4.1-a - Auditoría interna del proyecto

### Añadido
- Nueva opción `7. Auditoría del proyecto` en el menú principal.
- Nuevo módulo `modules/auditoria.py`.
- Carpeta automática `Resultados_Auditoria/`.
- Informe Word de auditoría.
- Informe Excel de auditoría.
- Resumen TXT de auditoría.

### Mantenido
- No se modifica el flujo de análisis Kubios OCR.
- No se modifica la gestión de participantes.
- No se modifica la validación científica SVF.

# CHANGELOG

## v2.1.0 - 2026-08-14

- Integración de la arquitectura modular desarrollada en la rama local v4.3.
- Añadidos módulos de participantes, datos públicos, validación y auditoría.
- Añadido cálculo de LnRMSSD y selección flexible del intervalo de fechas.
- Añadidas pruebas automatizadas y herramientas de distribución para Windows.
- Reforzada la exclusión de datos privados, resultados, cachés y binarios.
- Conservado `main_public.py` como lanzador seguro del entorno público.

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

## v1.4.1-a - Auditoría + tests básicos

### Añadido
- Carpeta `tests/` con pruebas automáticas básicas reales.
- Validación de estructura principal del proyecto.
- Validación de compilación de archivos Python principales.
- Validación básica del módulo de auditoría.

### Objetivo
- Elevar el proyecto desde auditoría avanzada hacia cierre estable, sin añadir nuevas funciones al análisis OCR.


## v1.4.6-a

### Fixed
- Sustituido el build anterior por un generador limpio con entorno virtual aislado.
- Evita arrastrar librerías globales no usadas como torch, tensorflow, torchvision, onnxruntime o pyarrow.
- El generador ya no usa requirements.txt completo para compilar el ejecutable.
