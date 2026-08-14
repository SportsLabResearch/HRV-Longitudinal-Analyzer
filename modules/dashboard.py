
from pathlib import Path
import platform, sys, ast

def mostrar_dashboard():
    root=Path(__file__).resolve().parents[1]
    py=list(root.rglob("*.py"))
    loc=0; funcs=0; classes=0
    for f in py:
        try:
            s=f.read_text(encoding="utf-8",errors="ignore")
            loc+=len(s.splitlines())
            t=ast.parse(s)
            funcs+=sum(isinstance(n,ast.FunctionDef) for n in ast.walk(t))
            classes+=sum(isinstance(n,ast.ClassDef) for n in ast.walk(t))
        except Exception:
            pass
    print("="*62)
    print(" HRV-LONGITUDINAL-ANALYZER - DASHBOARD")
    print("="*62)
    print(f"Sistema: {platform.system()} {platform.release()}")
    print(f"Python : {platform.python_version()}")
    print(f"Proyecto: {root}")
    print("-"*62)
    print(f"Archivos Python : {len(py)}")
    print(f"Líneas de código: {loc}")
    print(f"Funciones       : {funcs}")
    print(f"Clases          : {classes}")
    print("-"*62)
    for d in ["Datos","Resultados","modules","docs","tests","legacy"]:
        print(("✔" if (root/d).exists() else "✖"), d)
    print("="*62)
