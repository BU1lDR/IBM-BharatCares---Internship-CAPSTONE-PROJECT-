"""Execute the capstone notebook in place and report any cell failure.

Runs every cell headlessly with the repository root as the working directory,
writes the executed notebook back over itself, then exits non-zero if any cell
produced an error output - so a broken run cannot pass unnoticed.

Usage
-----
    python tools/run_notebook.py

Equivalent to Kernel > Restart Kernel and Run All Cells in Jupyter. Requires
nbclient (installed with the `notebook` package).
"""
import sys
from pathlib import Path

import nbformat
from nbclient import NotebookClient

ROOT = Path(__file__).resolve().parents[1]
NB = ROOT / "AryanVerma_RetailCustomerSegmentationAnalysis.ipynb"

nb = nbformat.read(NB, as_version=4)
client = NotebookClient(nb, timeout=2400, kernel_name="python3",
                        allow_errors=True,
                        resources={"metadata": {"path": str(ROOT)}})
print(f"executing {NB.name} ({len(nb.cells)} cells) ...", flush=True)
client.execute()
nbformat.write(nb, NB)

failed = 0
for i, cell in enumerate(nb.cells):
    for out in cell.get("outputs", []):
        if out.get("output_type") == "error":
            failed += 1
            print(f"\n{'=' * 70}\nCELL {i} FAILED: {out.get('ename')}: {out.get('evalue')}")
            print("--- source ---")
            print(cell["source"][:1400])
            print("--- traceback ---")
            print("\n".join(out.get("traceback", []))[-2500:])

n_code = sum(1 for c in nb.cells if c.cell_type == "code")
print(f"\n{'=' * 70}\n{n_code} code cells executed, {failed} raised an error.")
sys.exit(1 if failed else 0)
