import json
import unittest
from pathlib import Path


class NotebookTests(unittest.TestCase):
    def test_code_cells_compile(self):
        root = Path(__file__).resolve().parents[1]
        path = root / "Raportare_OMEC_3019_2025.ipynb"
        notebook = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(notebook["nbformat"], 4)
        for index, cell in enumerate(notebook["cells"], start=1):
            if cell["cell_type"] == "code":
                compile("".join(cell["source"]), f"cell-{index}", "exec")
                self.assertFalse(cell.get("outputs"))


if __name__ == "__main__":
    unittest.main()

