# -*- coding: utf-8 -*-
"""Lanzador público de HRV Longitudinal Analyzer."""

from __future__ import annotations

import re
import traceback
from pathlib import Path

import main as core

PROJECT_VERSION = "v2.1.0"

core.RESULTADOS_DIR_NAME = "Resultados_publicos"
core.INFORMES_DIR = core.SCRIPT_DIR / "Resultados_publicos"


def resolver_directorio_publico(root_dir):
    return Path(root_dir).resolve() / "Datos_publicos"


def listar_sujetos_publicos(root_dir):
    root_dir = Path(root_dir).resolve()

    if not root_dir.exists():
        return []

    sujetos = [
        carpeta
        for carpeta in root_dir.iterdir()
        if carpeta.is_dir()
        and re.fullmatch(
            r"Sujeto[_\-\s]*\d+",
            carpeta.name.strip(),
            flags=re.IGNORECASE,
        )
    ]

    return sorted(
        sujetos,
        key=lambda carpeta: int(re.search(r"\d+", carpeta.name).group()),
    )


def sin_seleccion_activa(*args, **kwargs):
    return []


def sin_preanalisis_privado(*args, **kwargs):
    return None


core.resolver_directorio_participantes = resolver_directorio_publico
core.listar_carpetas_participantes = listar_sujetos_publicos
core.selected_paths_from_active_selection = sin_seleccion_activa
core.mostrar_preanalisis_seleccion_activa = sin_preanalisis_privado


def main() -> None:
    while True:
        print("\n" + "=" * 60)
        print(f" HRV-LONGITUDINAL-ANALYZER {PROJECT_VERSION}")
        print(" SportsLabResearch")
        print("=" * 60)
        print("1. Analizar imágenes de Kubios")
        print("0. Salir")

        opcion = input("\nSelecciona una opción: ").strip()

        if opcion == "1":
            core.ejecutar_analisis()
        elif opcion == "0":
            print("Saliendo...")
            break
        else:
            print("Opción no válida.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print("\nERROR DURANTE LA EJECUCIÓN")
        print(str(exc))
        print(traceback.format_exc())
        input("Pulsa ENTER para salir...")
