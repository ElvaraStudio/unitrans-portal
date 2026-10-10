"""Generate a tiny workbook + PDF folder for smoke-testing the built exe.

Not part of the app. Run via ``smoke_test.py``, which uses this to prove the
packaged executable still performs a real sort end to end.
"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

# Container numbers are valid ISO 6346 check-digit values so the app's
# container regex and checksum logic accept them.
CONTAINERS = ["XAPU1814169", "TCLU1234562", "MSCU7654323"]

SEQ_HEADER = "\u5e8f\u53f7"
CONTAINER_HEADER = "\u7bb1\u53f7"


def make_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append([SEQ_HEADER, CONTAINER_HEADER])
    for index, container in enumerate(CONTAINERS, start=1):
        sheet.append([index, container])
    workbook.save(path)


def make_pdf(path: Path, label: str, pages: int = 2) -> None:
    pdf = canvas.Canvas(str(path), pagesize=A4)
    for page in range(1, pages + 1):
        pdf.setFont("Helvetica", 18)
        pdf.drawString(72, 760, f"Waybill {label} - page {page}")
        # Form number at the top-left, as the real waybills carry it.
        pdf.setFont("Helvetica", 10)
        pdf.drawString(72, 800, "1")
        pdf.showPage()
    pdf.save()


def main(argv: list[str]) -> int:
    target = Path(argv[0])
    workbook_dir = target / "workbook"
    pdf_dir = target / "pdfs"
    workbook_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)

    make_workbook(workbook_dir / "waybills.xlsx")
    for container in CONTAINERS:
        make_pdf(pdf_dir / f"{container}.pdf", container)

    print(f"workbook: {workbook_dir / 'waybills.xlsx'}")
    print(f"pdfs    : {pdf_dir} ({len(CONTAINERS)} files)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
