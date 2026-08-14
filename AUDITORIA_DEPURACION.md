# Auditoría de depuración previa a publicación

Fecha: 14 de agosto de 2026.

## Resultado

- Código Python compilado correctamente.
- No se detectaron nombres del mapa privado de participantes en el candidato.
- No se detectaron credenciales mediante la búsqueda de patrones habituales.
- Se excluyeron datos originales, datos seudonimizados, correspondencias privadas,
  resultados generados, cachés y paquetes previos.
- El generador de distribuciones se corrigió para impedir que esas carpetas se
  incorporen accidentalmente a futuros ZIP.
- El lanzador público se renombró de `main.pycls` a `main_publico.py` para que su
  función y tipo de archivo sean reconocibles.

## Incidencias pendientes antes de publicar una versión estable

- El archivo `LICENSE` está vacío; debe seleccionarse y documentarse una licencia.
- La comparación con la versión pública `v2.0.0` se realizó antes de preparar
  este candidato `v2.1.0`.
