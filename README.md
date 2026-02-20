# CritPathGenerator

A simple Python utility to compute the Critical Path Method (CPM) for project schedules and generate a PDF report including a network diagram.

## Features

- Reads tasks from an Excel file (`input/project.xlsx` by default)
- Calculates earliest and latest start/finish times, slack, and critical path
- Produces a tabular report and network diagram in PDF format

## Requirements

- Python 3.8+
- Dependencies listed in `requirements.txt` (install via `pip install -r requirements.txt`)

## Usage

1. Populate `input/project.xlsx` with columns: `Task`, `Duration`, `Predecessors`.
2. Run the script:

   ```bash
   python main.py
   ```

   (or import and call `compute_cpm_and_export_pdf` from another script)

3. Open `output/CPM_Report.pdf` to view the results.

## Project Structure

```
main.py                # core logic
input/                 # Excel input files
output/                # generated reports and diagrams
network_temp/          # temporary graph files (ignored)
requirements.txt       # dependencies
.gitignore
README.md
```

## Notes

- Modify the paths in `main.py` if you place files elsewhere.
- The network diagram is generated using Graphviz; ensure it is installed and on your PATH.

Enjoy using CritPathGenerator!"}