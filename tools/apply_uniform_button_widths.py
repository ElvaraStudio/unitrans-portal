"""One-off: give action buttons a shared, language-stable width.

Applies ``label_key=`` to every action-button call site so the button factory
can assign one width that fits the widest label in *any* language. Without it
each button sizes to its own text, so buttons within a column differ and the
whole column resizes when the language changes.

Run from the project root:

    python tools/apply_uniform_button_widths.py --dry-run
    python tools/apply_uniform_button_widths.py
"""

from __future__ import annotations

import argparse
import ast
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TARGET = ROOT / "waybill_sorter.py"

#: (exact existing fragment, replacement). Each must match exactly once.
EDITS: list[tuple[str, str]] = [
    # --- workflow selector: cards share one button width, stretched to fill ---
    (
        'self._primary_button(shell, title, command).grid(row=0, column=1, rowspan=2, sticky="e", padx=18)',
        'self._primary_button(shell, title, command, label_key=button_key).grid(\n'
        '            row=0, column=1, rowspan=2, sticky="nsew", padx=18, pady=16\n'
        '        )',
    ),
    # --- unlock screen ---
    (
        'self.unlock_button = self._primary_button(action_stack, self.tr("unlock_pdfs"), self.start_unlock_pdfs)',
        'self.unlock_button = self._primary_button(\n'
        '            action_stack, self.tr("unlock_pdfs"), self.start_unlock_pdfs, label_key="unlock_pdfs"\n'
        '        )',
    ),
    (
        'self.unlock_open_output_button = self._primary_button(action_stack, self.tr("open_output"), self.open_unlock_output)',
        'self.unlock_open_output_button = self._primary_button(\n'
        '            action_stack, self.tr("open_output"), self.open_unlock_output, label_key="open_output"\n'
        '        )',
    ),
    # --- driver screen ---
    (
        'self.driver_fill_button = self._primary_button(action_stack, self.tr("fill_driver_numbers"), self.start_driver_fill)',
        'self.driver_fill_button = self._primary_button(\n'
        '            action_stack, self.tr("fill_driver_numbers"), self.start_driver_fill,\n'
        '            label_key="fill_driver_numbers",\n'
        '        )',
    ),
    (
        'self.driver_open_output_button = self._primary_button(action_stack, self.tr("open_output"), self.open_driver_output)',
        'self.driver_open_output_button = self._primary_button(\n'
        '            action_stack, self.tr("open_output"), self.open_driver_output, label_key="open_output"\n'
        '        )',
    ),
    # --- sort screen ---
    (
        'self.sort_button = self._primary_button(action_stack, self.tr("sort"), self.start_sort)',
        'self.sort_button = self._primary_button(\n'
        '            action_stack, self.tr("sort"), self.start_sort, label_key="sort"\n'
        '        )',
    ),
    (
        'self.make_pdf_button = self._primary_button(action_stack, self.tr("make_pdf"), self.create_print_pdf)',
        'self.make_pdf_button = self._primary_button(\n'
        '            action_stack, self.tr("make_pdf"), self.create_print_pdf, label_key="make_pdf"\n'
        '        )',
    ),
    (
        'self.open_output_button = self._primary_button(action_stack, self.tr("open_output"), self.open_output)',
        'self.open_output_button = self._primary_button(\n'
        '            action_stack, self.tr("open_output"), self.open_output, label_key="open_output"\n'
        '        )',
    ),
    (
        'self.options_button = self._secondary_button(action_stack, self.tr("options"), self.open_options)',
        'self.options_button = self._secondary_button(\n'
        '            action_stack, self.tr("options"), self.open_options, label_key="options"\n'
        '        )',
    ),
    (
        'self._secondary_button(excel_body, self.tr("excel_columns"), self.open_excel_columns_dialog).grid(',
        'self._secondary_button(\n'
        '            excel_body, self.tr("excel_columns"), self.open_excel_columns_dialog,\n'
        '            label_key="excel_columns",\n'
        '        ).grid(',
    ),
    # --- dialogs ---
    (
        'close_button = self._primary_button(body, self.tr("close_options"), window.destroy)',
        'close_button = self._primary_button(\n'
        '            body, self.tr("close_options"), window.destroy, label_key="close_options"\n'
        '        )',
    ),
    (
        'check_button = self._primary_button(body, self.tr("settings_check_now"), lambda: None)',
        'check_button = self._primary_button(\n'
        '            body, self.tr("settings_check_now"), lambda: None, label_key="settings_check_now"\n'
        '        )',
    ),
    (
        'close_button = self._secondary_button(body, self.tr("settings_close"), window.destroy)',
        'close_button = self._secondary_button(\n'
        '            body, self.tr("settings_close"), window.destroy, label_key="settings_close"\n'
        '        )',
    ),
    (
        'self._primary_button(buttons, self.tr("update_download"), open_download).pack(side=tk.RIGHT)',
        'self._primary_button(\n'
        '            buttons, self.tr("update_download"), open_download, label_key="update_download"\n'
        '        ).pack(side=tk.RIGHT)',
    ),
    (
        'self._secondary_button(buttons, self.tr("cancel"), window.destroy).pack(side=tk.RIGHT, padx=(8, 0))',
        'self._secondary_button(\n'
        '            buttons, self.tr("cancel"), window.destroy, label_key="cancel"\n'
        '        ).pack(side=tk.RIGHT, padx=(8, 0))',
    ),
]

#: The three "Clear Log" secondary buttons, one per screen. Handled together
#: because the surrounding code differs slightly on each screen.
CLEAR_LOG_OLD = 'self._secondary_button(action_stack, self.tr("clear_log"), self.clear_log).pack(fill=tk.X)'
CLEAR_LOG_NEW = (
    'self._secondary_button(\n'
    '            action_stack, self.tr("clear_log"), self.clear_log, label_key="clear_log"\n'
    '        ).pack(fill=tk.X)'
)


def apply(path: Path, dry_run: bool = False) -> int:
    source = path.read_text(encoding="utf-8")
    ast.parse(source)  # must be valid before we touch it

    applied, skipped = 0, []
    for old, new in EDITS:
        count = source.count(old)
        if count != 1:
            skipped.append((old[:64], count))
            continue
        source = source.replace(old, new, 1)
        applied += 1

    clear_log_count = source.count(CLEAR_LOG_OLD)
    if clear_log_count >= 1:
        source = source.replace(CLEAR_LOG_OLD, CLEAR_LOG_NEW)
        applied += clear_log_count
    else:
        skipped.append((CLEAR_LOG_OLD[:64], clear_log_count))

    if skipped:
        print("Refusing to write - these fragments did not match exactly once:")
        for fragment, count in skipped:
            print(f"  ({count} matches) {fragment}")
        return 1

    ast.parse(source)
    if dry_run:
        print(f"Dry run: would apply {applied} edits to {path.name}.")
        return 0

    path.write_text(source, encoding="utf-8")
    print(f"Applied {applied} edits to {path.name}.")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args(argv)
    return apply(TARGET, dry_run=args.dry_run)


if __name__ == "__main__":
    sys.exit(main())
