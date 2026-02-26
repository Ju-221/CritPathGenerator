import shutil
import subprocess
import sys
from pathlib import Path


def test_main_generates_pdf_report():
    repo_root = Path(__file__).resolve().parents[1]
    input_dir = repo_root / "input"
    output_dir = repo_root / "output"

    sample_input = input_dir / "project.xlsx.example"
    expected_input = input_dir / "project.xlsx"
    report_file = output_dir / "CPM_Report.pdf"
    network_image = output_dir / "network_diagram.png"

    assert sample_input.exists(), "Sample input workbook is missing: input/project.xlsx.example"

    original_input_bytes = expected_input.read_bytes() if expected_input.exists() else None

    shutil.copyfile(sample_input, expected_input)

    if report_file.exists():
        report_file.unlink()
    if network_image.exists():
        network_image.unlink()

    try:
        result = subprocess.run(
            [sys.executable, "main.py"],
            cwd=repo_root,
            capture_output=True,
            text=True,
            check=False,
        )

        assert result.returncode == 0, (
            "main.py failed\n"
            f"stdout:\n{result.stdout}\n"
            f"stderr:\n{result.stderr}"
        )
        assert report_file.exists(), "Expected output/CPM_Report.pdf to be generated"
        assert report_file.stat().st_size > 0, "Generated CPM_Report.pdf is empty"
    finally:
        if report_file.exists():
            report_file.unlink()
        if network_image.exists():
            network_image.unlink()

        if original_input_bytes is None:
            if expected_input.exists():
                expected_input.unlink()
        else:
            expected_input.write_bytes(original_input_bytes)
