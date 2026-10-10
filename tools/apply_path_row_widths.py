"""One-off: give the file-picker buttons in the path rows a shared width.

The "Choose Excel" / "Choose PDF Folder" / "Choose Image Folder" buttons sit at
the top of each section. Each sized to its own label, so they differed between
sections and resized when the language changed - which moved the row and the
entry box next to it.

Run from the project root.
"""

from __future__ import annotations

import ast
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TARGET = ROOT / "waybill_sorter.py"

EDITS: list[tuple[str, str]] = [
    (
        'self._path_row(folder_body, self.tr("choose_folder"), self.choose_unlock_pdf_folder, self.unlock_folder_path)',
        'self._path_row(\n'
        '            folder_body, self.tr("choose_folder"), self.choose_unlock_pdf_folder,\n'
        '            self.unlock_folder_path, label_key="choose_folder",\n'
        '        )',
    ),
    (
        'self._path_row(excel_body, self.tr("choose_excel"), self.choose_driver_excel, self.driver_excel_path)',
        'self._path_row(\n'
        '            excel_body, self.tr("choose_excel"), self.choose_driver_excel,\n'
        '            self.driver_excel_path, label_key="choose_excel",\n'
        '        )',
    ),
    (
        'self._path_row(image_body, self.tr("choose_images"), self.choose_driver_image_folder, self.driver_image_folder)',
        'self._path_row(\n'
        '            image_body, self.tr("choose_images"), self.choose_driver_image_folder,\n'
        '            self.driver_image_folder, label_key="choose_images",\n'
        '        )',
    ),
    (
        'self._path_row(excel_body, self.tr("choose_excel"), self.choose_excel, self.excel_path)',
        'self._path_row(\n'
        '            excel_body, self.tr("choose_excel"), self.choose_excel,\n'
        '            self.excel_path, label_key="choose_excel",\n'
        '        )',
    ),
    (
        'self._path_row(pdf_body, self.tr("choose_folder"), self.choose_pdf_folder, self.pdf_folder)',
        'self._path_row(\n'
        '            pdf_body, self.tr("choose_folder"), self.choose_pdf_folder,\n'
        '            self.pdf_folder, label_key="choose_folder",\n'
        '        )',
    ),
]


def main() -> int:
    source = TARGET.read_text(encoding="utf-8")
    ast.parse(source)

    applied, problems = 0, []
    for old, new in EDITS:
        count = source.count(old)
        if count != 1:
            problems.append((count, old[:70]))
            continue
        source = source.replace(old, new, 1)
        applied += 1

    if problems:
        print("Refusing to write - these fragments did not match exactly once:")
        for count, fragment in problems:
            print(f"  ({count} matches) {fragment}")
        return 1

    ast.parse(source)
    TARGET.write_text(source, encoding="utf-8")
    print(f"Applied {applied} path-row edits to {TARGET.name}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
