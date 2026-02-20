# CritPathGenerator

Honestly screw making CPM manually, here's a python script for that.

A simple Python utility to compute the Critical Path Method (CPM) for project schedules and generate a PDF report including a network diagram.

This program follows CPM terminology from project management. When multiple critical paths exist, it selects one to highlight based on:

1. **Total duration** (longest time to complete),
2. **Node count** as a tiebreaker (more tasks preferred).

Only the chosen path is colored red in the diagram.

## Features

- Reads tasks from an Excel file (`input/project.xlsx` by default, see porject.xlsx.example)
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
input/                 # Excel input files - rename project.xlsx.example to see the expected input
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