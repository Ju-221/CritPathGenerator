import os
import pandas as pd
from collections import defaultdict, deque
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Table, TableStyle, PageBreak, Image
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from graphviz import Digraph


import shutil


def _ensure_graphviz_installed():
    """Check that the Graphviz `dot` executable is available on the PATH.

    Raises:
        RuntimeError: if `dot` cannot be found.
    """
    if shutil.which("dot") is None:
        raise RuntimeError(
            "Graphviz is not installed or 'dot' is not on your PATH.\n"
            "Please install Graphviz from https://graphviz.org/download/ "
            "and ensure the installation directory is added to your PATH."
        )


def compute_cpm_and_export_pdf(input_path, output_path):
    # verify external dependency before doing any work
    _ensure_graphviz_installed()

    # -------------------------
    # Ensure output directory exists
    # -------------------------
    output_dir = os.path.dirname(output_path)

    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # -------------------------
    # Load Excel
    # -------------------------
    df = pd.read_excel(input_path)
    df['Predecessors'] = df['Predecessors'].fillna('').astype(str)

    tasks = df['Task'].tolist()
    duration = dict(zip(df['Task'], df['Duration']))

    predecessors = {}
    successors = defaultdict(list)

    for _, row in df.iterrows():
        task = row['Task']
        preds = [p.strip() for p in row['Predecessors'].split(',') if p.strip()]
        predecessors[task] = preds
        for p in preds:
            successors[p].append(task)

    # -------------------------
    # Topological Sort
    # -------------------------
    in_degree = {t: len(predecessors[t]) for t in tasks}
    queue = deque([t for t in tasks if in_degree[t] == 0])
    topo_order = []

    while queue:
        node = queue.popleft()
        topo_order.append(node)
        for succ in successors[node]:
            in_degree[succ] -= 1
            if in_degree[succ] == 0:
                queue.append(succ)

    # -------------------------
    # Forward Pass
    # -------------------------
    ES, EF = {}, {}
    for task in topo_order:
        if not predecessors[task]:
            ES[task] = 0
        else:
            ES[task] = max(EF[p] for p in predecessors[task])
        EF[task] = ES[task] + duration[task]

    project_duration = max(EF.values())

    # -------------------------
    # Backward Pass
    # -------------------------
    LS, LF = {}, {}
    for task in reversed(topo_order):
        if not successors[task]:
            LF[task] = project_duration
        else:
            LF[task] = min(LS[s] for s in successors[task])
        LS[task] = LF[task] - duration[task]

    slack = {t: LS[t] - ES[t] for t in tasks}

    # determine a single critical path to highlight
    # we use two criteria in order:
    # 1. highest total duration (should equal project_duration for true critical paths)
    # 2. most nodes when durations tie
    # dynamic programming on topo_order: track best path info to each task
    # path_info: (total_duration, node_count, path_list)
    best_info = {}
    for task in topo_order:
        dur = duration[task]
        if not predecessors[task]:
            best_info[task] = (dur, 1, [task])
        else:
            # pick predecessor with maximum (total_duration, node_count)
            prev = max(
                predecessors[task],
                key=lambda p: best_info.get(p, (0, 0, []))[:2]
            )
            prev_dur, prev_count, prev_path = best_info.get(prev, (0, 0, []))
            best_info[task] = (prev_dur + dur, prev_count + 1, prev_path + [task])
    # evaluate terminal tasks to select the best path
    end_tasks = [t for t in tasks if not successors[t]]
    critical_path = []
    best_tuple = (0, 0)  # (total_duration, node_count)
    for t in end_tasks:
        tot_dur, cnt, path = best_info.get(t, (0, 0, []))
        candidate = (tot_dur, cnt)
        if candidate > best_tuple:
            best_tuple = candidate
            critical_path = path

    critical = critical_path

    # -------------------------
    # Generate Network Diagram
    # -------------------------
    dot = Digraph(format="png")
    dot.attr(rankdir="LR")

    critical_set = set(critical)
    for task in tasks:
        label = (
            f"{task}\n"
            f"-----------------\n"
            f"ES: {ES[task]} | EF: {EF[task]}\n"
            f"LS: {LS[task]} | LF: {LF[task]}\n"
            f"Slack: {slack[task]}"
        )

        # highlight only the selected critical path in red
        color = "red" if task in critical_set else "lightblue"

        dot.node(
            task,
            label=label,
            shape="box",
            style="filled",
            fillcolor=color
        )

    for task in tasks:
        for pred in predecessors[task]:
            dot.edge(pred, task)

    network_image_base = os.path.join(output_dir, "network_diagram")
    dot.render(network_image_base, cleanup=True)

    # -------------------------
    # Create PDF with mixed orientations
    # -------------------------
    from reportlab.lib.pagesizes import letter, landscape, portrait
    from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, NextPageTemplate

    # define page sizes
    portrait_size = portrait(letter)
    landscape_size = landscape(letter)

    # create frames covering the entire page for each orientation
    portrait_frame = Frame(
        0, 0, portrait_size[0], portrait_size[1], id='portrait')
    landscape_frame = Frame(
        0, 0, landscape_size[0], landscape_size[1], id='landscape')
    
    doc = BaseDocTemplate(output_path, pagesize=portrait_size)
    doc.addPageTemplates([
        PageTemplate(id='Portrait', frames=[portrait_frame], pagesize=portrait_size),
        PageTemplate(id='Landscape', frames=[landscape_frame], pagesize=landscape_size),
    ])
    
    elements = []
    styles = getSampleStyleSheet()

    elements.append(Paragraph("<b>CPM Project Report</b>", styles['Title']))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"Project Duration: {project_duration}", styles['Normal']))
    elements.append(Paragraph(f"Critical Path: {' → '.join(critical)}", styles['Normal']))
    elements.append(Spacer(1, 20))

    # include the raw input data from the Excel file
    elements.append(Paragraph("<b>Input Task Data</b>", styles['Heading2']))
    input_table_data = [df.columns.tolist()] + df[['Task','Duration','Predecessors']].values.tolist()
    input_table = Table(input_table_data)
    input_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    elements.append(input_table)
    elements.append(Spacer(1, 20))

    result = pd.DataFrame({
        "Task": tasks,
        "Duration": [duration[t] for t in tasks],
        "ES": [ES[t] for t in tasks],
        "EF": [EF[t] for t in tasks],
        "LS": [LS[t] for t in tasks],
        "LF": [LF[t] for t in tasks],
        "Slack": [slack[t] for t in tasks]
    })

    table_data = [result.columns.tolist()] + result.values.tolist()

    table = Table(table_data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    elements.append(table)

    # -------- Page Break (switch to landscape) --------
    elements.append(NextPageTemplate('Landscape'))
    elements.append(PageBreak())
    elements.append(Paragraph("<b>CPM Network Diagram</b>", styles['Title']))
    elements.append(Spacer(1, 20))

    elements.append(Image(network_image_base + ".png", width=500, height=350))

    doc.build(elements)

    print(f"PDF Report generated at: {output_path}")


# -------------------------
# Example usage
# -------------------------
compute_cpm_and_export_pdf(
    r"input/project.xlsx",
    r"output/CPM_Report.pdf"
)
