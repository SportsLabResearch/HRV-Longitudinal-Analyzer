# -*- coding: utf-8 -*-
"""
Instalador simple del módulo de auditoría.
Ejecutar desde la raíz del proyecto:

py scripts/instalar_auditoria.py

Hace copia de seguridad de main.py antes de modificarlo.
"""

from pathlib import Path
from datetime import datetime

ROOT = Path.cwd()
MAIN = ROOT / "main.py"

IMPORT_LINE = "from modules.auditoria import ejecutar_auditoria_proyecto\n"
MENU_LINE = "7. Auditoría del proyecto"
OPTION_LINE = "ejecutar_auditoria_proyecto()"


def backup(path: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = path.with_name(f"{path.stem}_backup_{stamp}{path.suffix}")
    dst.write_text(path.read_text(encoding="utf-8"), encoding="utf-8")
    return dst


def main():
    if not MAIN.exists():
        print("ERROR: no se encuentra main.py. Ejecuta este instalador desde la raíz del proyecto.")
        return

    text = MAIN.read_text(encoding="utf-8")
    bck = backup(MAIN)

    changed = False

    if "ejecutar_auditoria_proyecto" not in text:
        lines = text.splitlines(True)
        insert_at = 0
        for i, line in enumerate(lines):
            if line.startswith("import ") or line.startswith("from "):
                insert_at = i + 1
        lines.insert(insert_at, IMPORT_LINE)
        text = "".join(lines)
        changed = True

    if MENU_LINE not in text:
        text = text.replace("6. Validación científica (SVF)", "6. Validación científica (SVF)\n7. Auditoría del proyecto")
        changed = True

    # Inserción conservadora: añade caso 7 cerca del caso 6 si reconoce estructuras típicas.
    if OPTION_LINE not in text:
        candidates = [
            ('elif opcion == "6":', 'elif opcion == "7":\n            ejecutar_auditoria_proyecto()\n'),
            ("elif opcion == '6':", "elif opcion == '7':\n            ejecutar_auditoria_proyecto()\n"),
            ('if opcion == "6":', 'elif opcion == "7":\n            ejecutar_auditoria_proyecto()\n'),
            ("if opcion == '6':", "elif opcion == '7':\n            ejecutar_auditoria_proyecto()\n"),
        ]
        inserted = False
        for marker, block in candidates:
            pos = text.find(marker)
            if pos != -1:
                next_zero = text.find('elif opcion == "0"', pos)
                if next_zero == -1:
                    next_zero = text.find("elif opcion == '0'", pos)
                if next_zero != -1:
                    text = text[:next_zero] + block + text[next_zero:]
                    inserted = True
                    changed = True
                    break
        if not inserted:
            print("AVISO: no pude insertar automáticamente la opción 7 en la lógica del menú.")
            print("Añade manualmente: elif opcion == '7': ejecutar_auditoria_proyecto()")

    if changed:
        MAIN.write_text(text, encoding="utf-8")
        print("Instalación aplicada.")
    else:
        print("No había cambios que aplicar.")

    print(f"Copia de seguridad: {bck.name}")
    print("Ahora ejecuta: py main.py")


if __name__ == "__main__":
    main()
