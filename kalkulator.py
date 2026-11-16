import ast, io, math, builtins, contextlib
import numpy as np
import pandas as pd

TERLARANG = {"open", "exec", "eval", "__import__", "compile",
             "input", "globals", "locals", "getattr", "setattr", "breakpoint"}

BUILTIN_AMAN = {n: getattr(builtins, n) for n in
                ["len", "range", "sum", "min", "max", "round", "abs", "print",
                 "float", "int", "str", "list", "dict", "sorted", "enumerate", "zip"]}

def periksa(kode: str) -> None:
    for node in ast.walk(ast.parse(kode)):
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            raise ValueError("import tidak diizinkan; pd, np, dan math sudah tersedia")
        if isinstance(node, ast.Attribute) and node.attr.startswith("_"):
            raise ValueError("akses atribut privat tidak diizinkan")
        if isinstance(node, ast.Name) and node.id in TERLARANG:
            raise ValueError(f"nama '{node.id}' tidak diizinkan")

def jalankan(kode: str, data_sesi: dict) -> dict:
    try:
        periksa(kode)
        lingkungan = {"__builtins__": BUILTIN_AMAN, "pd": pd, "np": np,
                      "math": math, "data": data_sesi}
        keluaran = io.StringIO()
        with contextlib.redirect_stdout(keluaran):
            exec(kode, lingkungan)
        return {"ok": True, "output": keluaran.getvalue()[:4000],
                "hasil": str(lingkungan.get("hasil"))}
    except Exception as e:
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}