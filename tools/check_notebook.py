"""Validează structura JSON și sintaxa celulelor Python ale notebookului."""

from __future__ import annotations

import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
path = ROOT / "Raportare_OMEC_3019_2025.ipynb"
notebook = json.loads(path.read_text(encoding="utf-8"))

if notebook.get("nbformat") != 4:
    raise AssertionError("Notebookul trebuie să folosească nbformat 4.")

code_cells = 0
for index, cell in enumerate(notebook.get("cells", []), start=1):
    if cell.get("cell_type") != "code":
        continue
    code_cells += 1
    compile("".join(cell.get("source", [])), f"cell-{index}", "exec")
    if cell.get("outputs"):
        raise AssertionError(f"Celula {index} conține output salvat.")

if code_cells == 0:
    raise AssertionError("Notebookul nu conține celule Python.")

print(
    f"Notebook valid: {len(notebook['cells'])} celule, "
    f"{code_cells} celule Python, fără outputuri."
)

