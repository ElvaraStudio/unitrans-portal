from __future__ import annotations

import argparse
import csv
import datetime as dt
import hashlib
import io
import json
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import unicodedata
import webbrowser
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

try:
    from openpyxl import load_workbook
    import msoffcrypto
    from pypdf import PdfReader, PdfWriter
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfgen import canvas
    from PIL import Image, ImageDraw, ImageFont, ImageTk
except ImportError as exc:  # pragma: no cover - shown only when run from source without deps.
    load_workbook = None
    PdfReader = None
    PdfWriter = None
    ImageReader = None
    Image = None
    ImageDraw = None
    ImageFont = None
    ImageTk = None
    IMPORT_ERROR = exc
else:
    IMPORT_ERROR = None

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import settings as app_settings
import theme
from theme import DARK, LIGHT
from updater import (
    UpdateCancelled,
    UpdateDownloadError,
    check_for_update,
    download_update,
    format_size,
    launch_update_helper,
    supports_self_update,
)
from version import APP_VERSION, display_version, version_label


# ---------------------------------------------------------------------------
# Colour tokens. Read from theme.py at startup and refreshed whenever the user
# switches between light and dark, so no widget needs to know which scheme is
# in use. Reassigning these and recolouring existing widgets applies a theme
# change - see apply_theme() and WaybillSorterApp.toggle_theme().
# ---------------------------------------------------------------------------
def _load_colour_tokens() -> dict[str, str]:
    """Current value of every colour token, keyed by its constant name."""
    return {
        "BG_WINDOW": theme.colour("window"),
        "BG_PANEL": theme.colour("surface_alt"),
        "BG_LIST": theme.colour("list_bg"),
        "BG_WARNING": theme.colour("danger_bg"),
        "SURFACE": theme.colour("surface"),
        "ON_ACCENT": theme.colour("on_accent"),
        # The brand mark is fixed navy artwork, so it always sits on a light
        # plate rather than following the scheme.
        "LIGHT_LOGO_PLATE": "#F4F6F9",
        "ACCENT": theme.colour("accent"),
        "ACCENT_HOVER": theme.colour("accent_hover"),
        "ACCENT_DISABLED": theme.colour("accent_disabled"),
        "ACCENT_SOFT": theme.colour("accent_soft"),
        "ACCENT_SOFT_HOVER": theme.colour("accent_soft_hover"),
        "ACCENT_TEXT": theme.colour("accent_text"),
        "TEXT": theme.colour("text"),
        "TEXT_STRONG": theme.colour("text_strong"),
        "TEXT_HEADING": theme.colour("text_heading"),
        "TEXT_MUTED": theme.colour("text_muted"),
        "TEXT_FAINT": theme.colour("text_faint"),
        "TEXT_LABEL": theme.colour("text_label"),
        "TEXT_ENTRY": theme.colour("text_entry"),
        "BORDER": theme.colour("border"),
        "DIVIDER": theme.colour("divider"),
        "BORDER_WARNING": theme.colour("warning_border"),
        "SCROLLBAR": theme.colour("scrollbar"),
        "SCROLLBAR_TROUGH": theme.colour("scrollbar_trough"),
        "STATE_SUCCESS": theme.colour("success"),
        "STATE_ERROR": theme.colour("error"),
        "STATE_WARNING": theme.colour("warning"),
    }


(
    BG_WINDOW, BG_PANEL, BG_LIST, BG_WARNING, SURFACE, ON_ACCENT,
    LIGHT_LOGO_PLATE,
    ACCENT, ACCENT_HOVER, ACCENT_DISABLED, ACCENT_SOFT, ACCENT_SOFT_HOVER,
    ACCENT_TEXT, TEXT, TEXT_STRONG, TEXT_HEADING, TEXT_MUTED, TEXT_FAINT,
    TEXT_LABEL, TEXT_ENTRY, BORDER, DIVIDER, BORDER_WARNING, SCROLLBAR,
    SCROLLBAR_TROUGH, STATE_SUCCESS, STATE_ERROR, STATE_WARNING,
) = _load_colour_tokens().values()


def apply_theme(name: str) -> None:
    """Switch the active colour scheme and refresh every token constant."""
    global BG_WINDOW, BG_PANEL, BG_LIST, BG_WARNING, SURFACE, ON_ACCENT
    global LIGHT_LOGO_PLATE
    global ACCENT, ACCENT_HOVER, ACCENT_DISABLED, ACCENT_SOFT, ACCENT_SOFT_HOVER
    global ACCENT_TEXT, TEXT, TEXT_STRONG, TEXT_HEADING, TEXT_MUTED, TEXT_FAINT
    global TEXT_LABEL, TEXT_ENTRY, BORDER, DIVIDER, BORDER_WARNING
    global SCROLLBAR, SCROLLBAR_TROUGH, STATE_SUCCESS, STATE_ERROR, STATE_WARNING

    theme.set_active(name)
    tokens = _load_colour_tokens()
    BG_WINDOW = tokens["BG_WINDOW"]
    BG_PANEL = tokens["BG_PANEL"]
    BG_LIST = tokens["BG_LIST"]
    BG_WARNING = tokens["BG_WARNING"]
    SURFACE = tokens["SURFACE"]
    ON_ACCENT = tokens["ON_ACCENT"]
    LIGHT_LOGO_PLATE = tokens["LIGHT_LOGO_PLATE"]
    ACCENT = tokens["ACCENT"]
    ACCENT_HOVER = tokens["ACCENT_HOVER"]
    ACCENT_DISABLED = tokens["ACCENT_DISABLED"]
    ACCENT_SOFT = tokens["ACCENT_SOFT"]
    ACCENT_SOFT_HOVER = tokens["ACCENT_SOFT_HOVER"]
    ACCENT_TEXT = tokens["ACCENT_TEXT"]
    TEXT = tokens["TEXT"]
    TEXT_STRONG = tokens["TEXT_STRONG"]
    TEXT_HEADING = tokens["TEXT_HEADING"]
    TEXT_MUTED = tokens["TEXT_MUTED"]
    TEXT_FAINT = tokens["TEXT_FAINT"]
    TEXT_LABEL = tokens["TEXT_LABEL"]
    TEXT_ENTRY = tokens["TEXT_ENTRY"]
    BORDER = tokens["BORDER"]
    DIVIDER = tokens["DIVIDER"]
    BORDER_WARNING = tokens["BORDER_WARNING"]
    SCROLLBAR = tokens["SCROLLBAR"]
    SCROLLBAR_TROUGH = tokens["SCROLLBAR_TROUGH"]
    STATE_SUCCESS = tokens["STATE_SUCCESS"]
    STATE_ERROR = tokens["STATE_ERROR"]
    STATE_WARNING = tokens["STATE_WARNING"]


SEQ_HEADER = "\u5e8f\u53f7"
CONTAINER_HEADER = "\u7bb1\u53f7"
CAR_NUMBER_HEADER = "\u8f66\u53f7"
CONTAINER_RE = re.compile(r"(?<![A-Z0-9])([A-Z]{4}\d{7})(?![A-Z0-9])")
CONTAINER_CHECK_VALUES = {
    **{str(number): number for number in range(10)},
    "A": 10,
    "B": 12,
    "C": 13,
    "D": 14,
    "E": 15,
    "F": 16,
    "G": 17,
    "H": 18,
    "I": 19,
    "J": 20,
    "K": 21,
    "L": 23,
    "M": 24,
    "N": 25,
    "O": 26,
    "P": 27,
    "Q": 28,
    "R": 29,
    "S": 30,
    "T": 31,
    "U": 32,
    "V": 34,
    "W": 35,
    "X": 36,
    "Y": 37,
    "Z": 38,
}
LOGO_FILE_NAME = "UniTrans Logo.png"
# APP_VERSION now lives in version.py so the build script and the app cannot
# disagree about which release this is.

#: Labels whose buttons are given one shared width, so they line up as a block
#: and do not resize when the language changes. Measured across every language,
#: not just the active one - see WaybillSorterApp._action_button_width.
ACTION_BUTTON_KEYS = frozenset({
    # Workflow selector
    "workflow_sort",
    "workflow_driver",
    "workflow_unlock",
    # Sort screen
    "sort",
    "make_pdf",
    "open_output",
    "options",
    "clear_log",
    "excel_columns",
    # Driver screen
    "fill_driver_numbers",
    # Unlock screen
    "unlock_pdfs",
    # Dialogs
    "back_to_workflows",
    "settings_check_now",
    "settings_close",
    "close_options",
    "apply",
    "cancel",
    "update_download",
    "update_later",
    # File-picker buttons in the path rows. These sit above a wide entry box,
    # so they are the same width as each other rather than as the tall action
    # buttons. Note the two distinct "folder" labels.
    "choose_excel",
    "choose_folder",
    "choose_unlock_files",
    "choose_unlock_folder",
    "choose_images",
})

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}
HEADER_ALIASES = {
    SEQ_HEADER: (SEQ_HEADER, "\u5e8f\u865f", "\u7f16\u53f7", "\u7de8\u865f"),
    CONTAINER_HEADER: (
        CONTAINER_HEADER,
        "\u7bb1\u865f",
        "\u67dc\u53f7",
        "\u67dc\u865f",
        "\u7bb1\u67dc\u53f7",
        "\u7bb1\u67dc\u865f",
        "\u96c6\u88c5\u7bb1\u53f7",
        "\u96c6\u88dd\u7bb1\u865f",
        "\u96c6\u88c5\u7bb1\u7bb1\u53f7",
        "\u96c6\u88dd\u7bb1\u7bb1\u865f",
        "\u96c6\u88c5\u7bb1\u53f7\u7801",
        "\u96c6\u88dd\u7bb1\u865f\u78bc",
        "container",
        "containerno",
        "containernumber",
    ),
    CAR_NUMBER_HEADER: (CAR_NUMBER_HEADER, "\u8eca\u865f", "\u8f66\u8f86", "\u8eca\u8f1b", "\u8f66\u8f86\u53f7", "\u8eca\u8f1b\u865f"),
}


@dataclass(frozen=True)
class WaybillRow:
    sequence: int
    container: str
    excel_row: int
    source_order: int = 0
    sequence_occurrence: int = 1
    sequence_note: str = ""


@dataclass(frozen=True)
class Token:
    chinese: str
    english: str
    aliases: tuple[str, ...]


@dataclass(frozen=True)
class LookupRule:
    number: int
    port_tokens: tuple[Token, ...]
    destination_tokens: tuple[Token, ...]
    english_rule: str


@dataclass
class PdfMatch:
    path: Path
    container: str | None
    method: str
    note: str


@dataclass
class SortResult:
    output_dir: Path
    sorted_dir: Path
    report_path: Path
    summary_path: Path
    expected_count: int
    pdf_count: int
    matched_count: int
    missing_count: int
    unmatched_count: int
    duplicate_count: int
    verified_copy_count: int
    verification_failure_count: int
    copy_amount: str
    copy_rule: str


@dataclass
class PrintBatchResult:
    path: Path
    waybill_count: int
    source_pages: int
    output_pages: int
    separator_pages: int


@dataclass
class PdfUnlockResult:
    output_dir: Path
    unlocked_dir: Path
    report_path: Path
    summary_path: Path
    total_count: int
    unlocked_count: int
    already_unlocked_count: int
    skipped_count: int
    failed_count: int


UnlockResult = PdfUnlockResult


@dataclass(frozen=True)
class DriverExcelRow:
    sequence: int
    excel_row: int
    container: str
    current_car_number: str


@dataclass(frozen=True)
class OcrTextResult:
    text: str
    confidence: float
    note: str = ""


@dataclass(frozen=True)
class RowMatch:
    row: DriverExcelRow | None
    method: str
    score: float = 0.0
    margin: float = 0.0
    note: str = ""


@dataclass
class DriverFillResult:
    output_dir: Path
    workbook_path: Path
    report_path: Path
    expected_count: int
    image_count: int
    filled_count: int
    already_filled_count: int
    unmatched_count: int
    review_count: int


@dataclass(frozen=True)
class PdfFolderAnalysis:
    pdf_count: int
    page_counts: tuple[int, ...]
    detected_supplementary_count: int | None


def token(chinese: str, english: str, aliases: Iterable[str] = ()) -> Token:
    return Token(chinese, english, tuple(aliases))


HORGOS = (
    token(
        "\u679c\u65af",
        "Horgos",
        (
            "\u970d\u679c\u65af",
            "\u970d\u5c14\u679c\u65af",
            "\u970d\u5c14\u679c\u65af\u53e3\u5cb8",
            "horgos",
            "khorgos",
        ),
    ),
    token(
        "\u970d\u5c14\u679c\u65af",
        "Horgos",
        ("\u679c\u65af", "\u970d\u679c\u65af", "horgos", "khorgos"),
    ),
)
TASHKENT_MUTU_BAKU = (
    token("\u5854\u4ec0\u5e72", "Tashkent", ("\u5854\u4ec0", "tashkent", "tash")),
    token("\u6728\u56fe", "Mutu", ("\u7a46\u56fe", "mutu")),
    token("\u5df4\u5e93", "Baku", ("baku",)),
    token("\u963f\u62c9\u6728\u56fe", "Almaty", ("\u963f\u62c9\u6728\u5716", "\u963f\u62c9\u6728", "almaty", "alma")),
)
POTI = (token("\u6ce2\u5b63", "Poti", ("\u6ce2\u8482", "poti")),)
GERMANY = (
    token("\u5fb7\u56fd", "Germany", ("germany", "german")),
    token("\u6c49\u5821", "Hamburg", ("hamburg", "hamb")),
    token("\u675c\u5821", "Duisburg", ("\u675c\u4f0a\u65af\u5821", "duisburg")),
    token(
        "\u9a6c\u62c9",
        "Malaszewicze",
        (
            "\u9a6c\u62c9\u820d\u7ef4\u5947",
            "\u9a6c\u62c9\u820d\u7ef4\u5207",
            "malaszewicze",
            "malaszewicz",
            "mala",
        ),
    ),
)
ALASHANKOU = (
    token(
        "\u5c71\u53e3",
        "Alashankou",
        ("\u963f\u62c9\u5c71\u53e3", "\u963f\u62c9", "alashankou", "ala"),
    ),
    token(
        "\u963f\u62c9\u5c71\u53e3",
        "Alashankou",
        ("\u5c71\u53e3", "\u963f\u62c9", "alashankou", "ala"),
    ),
)
EUROPE = (
    token("\u6b27\u6d32", "Europe", ("\u6b27\u7ebf", "europe", "eu")),
    token("Europe", "Europe", ("\u6b27\u6d32", "\u6b27\u7ebf", "eu")),
)
ASIA = (
    token("\u4e9a\u6d32", "Asia", ("\u4e9a\u7ebf", "asia")),
    token("Asia", "Asia", ("\u4e9a\u6d32", "\u4e9a\u7ebf")),
    token("\u54c8\u8428\u514b\u65af\u5766", "Kazakhstan", ("\u54c8\u8428\u514b", "kazakhstan", "kazakh")),
    token("\u5e93\u65af\u5854\u5948", "Kostanay", ("\u79d1\u65af\u5854\u5948", "kostanay", "kustanay", "kostanai")),
)
MANZHOULI = (token("\u6ee1\u6d32\u91cc", "Manzhouli", ("\u6ee1\u5dde\u91cc", "manzhouli", "mzl")),)
ERENHOT = (token("\u4e8c\u8fde", "Erenhot", ("\u4e8c\u8fde\u6d69\u7279", "erenhot", "erlian", "elht")),)


LOOKUP_RULES = (
    LookupRule(9, HORGOS, TASHKENT_MUTU_BAKU, "Horgos, destination is Tashkent / Mutu / Baku / Almaty"),
    LookupRule(9, (), TASHKENT_MUTU_BAKU, "Destination is Tashkent / Mutu / Baku / Almaty"),
    LookupRule(13, HORGOS, POTI, "Horgos, destination is Poti"),
    LookupRule(13, (), POTI, "Destination is Poti"),
    LookupRule(13, (), GERMANY, "Destination is Germany / Hamburg / Duisburg / Malaszewicze"),
    LookupRule(5, ALASHANKOU, EUROPE, "Alashankou, destination is Europe"),
    LookupRule(4, ALASHANKOU, ASIA, "Alashankou, destination is Asia"),
    LookupRule(2, MANZHOULI, (), "Manzhouli"),
    LookupRule(4, ERENHOT, (), "Erenhot"),
)


def normalize_lookup_text(text: str | None) -> str:
    if not text:
        return ""
    chars: list[str] = []
    for ch in text.lower():
        category = unicodedata.category(ch)
        if ch == "_" or category[0] in {"P", "S", "Z"}:
            continue
        chars.append(ch)
    return "".join(chars)


def token_aliases(item: Token) -> tuple[str, ...]:
    return (item.chinese, item.english, *item.aliases)


def first_token_match(text: str, tokens: tuple[Token, ...]) -> Token | None:
    normalized_text = normalize_lookup_text(text)
    for item in tokens:
        for alias in token_aliases(item):
            normalized_alias = normalize_lookup_text(alias)
            if normalized_alias and normalized_alias in normalized_text:
                return item
    return None


def lookup_copy_amount(document_name: str) -> tuple[str, str]:
    name = document_name.strip()
    if not name:
        return "-", "Nothing detected"

    for rule in LOOKUP_RULES:
        if rule.port_tokens and not first_token_match(name, rule.port_tokens):
            continue
        if rule.destination_tokens and not first_token_match(name, rule.destination_tokens):
            continue
        return str(rule.number), rule.english_rule

    return "-", "Nothing detected"


def excel_lookup_text(excel_path: Path, max_rows: int = 8, max_columns: int = 12) -> str:
    import_dependencies_ok()

    workbook = load_workbook(excel_path, data_only=True, read_only=True)
    try:
        worksheet = workbook.active
        parts = [excel_path.name]
        for row in worksheet.iter_rows(max_row=max_rows, max_col=max_columns, values_only=True):
            for value in row:
                text = "" if value is None else str(value).strip()
                if text:
                    parts.append(text)
        return " ".join(parts)
    finally:
        workbook.close()


def lookup_excel_copy_amount(excel_path: Path) -> tuple[str, str]:
    try:
        return lookup_copy_amount(excel_lookup_text(excel_path))
    except Exception:
        return lookup_copy_amount(excel_path.name)


def import_dependencies_ok() -> None:
    if IMPORT_ERROR is not None:
        raise RuntimeError(
            "Missing Python dependency. Install requirements.txt or run the packaged .exe. "
            f"Original error: {IMPORT_ERROR}"
        )


def _window_frame_size(widget: tk.Misc) -> tuple[int, int]:
    """Frame + border size Windows adds around a Tk window's client area.

    ``geometry()`` sets the client area, so sizing a window to fill a screen
    needs the non-client chrome added on top. Tk does not expose it, and
    ``winfo_rootx() - winfo_x()`` does not measure it (both are relative to the
    parent, which for a toplevel is the screen). ``AdjustWindowRectEx`` is the
    Win32 call that answers this exactly, including any DPI scaling.

    Returns ``(0, 0)`` when the metrics are unavailable, in which case the
    caller should leave the geometry alone.
    """
    try:
        import ctypes
        import ctypes.wintypes as wintypes

        widget.update_idletasks()
        user32 = ctypes.windll.user32

        # AdjustWindowRectEx works in physical pixels for the process DPI
        # awareness, while Tk's geometry() uses logical pixels. On a scaled
        # display the two differ, so scale using the window's own DPI.
        try:
            dpi = int(user32.GetDpiForWindow(int(widget.winfo_id())))
        except Exception:
            dpi = 96
        scale = (dpi / 96.0) if dpi > 0 else 1.0

        rect = wintypes.RECT(0, 0, 0, 0)
        # WS_OVERLAPPEDWINDOW: caption, thick frame, minimise/maximise boxes.
        WS_OVERLAPPEDWINDOW = 0x00CF0000
        if not user32.AdjustWindowRectEx(ctypes.byref(rect), WS_OVERLAPPEDWINDOW, False, 0):
            return (0, 0)

        width = int(round((rect.right - rect.left) * scale))
        height = int(round((rect.bottom - rect.top) * scale))
        if width <= 0 or height <= 0:
            return (0, 0)
        return (width, height)
    except Exception:
        return (0, 0)


def resource_path(file_name: str) -> Path:
    # _MEIPASS is where PyInstaller unpacks bundled data for a onefile build.
    # Path(__file__) covers running from source. The old third candidate pointed
    # at a branding folder on a different developer machine; it is gone.
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    candidates = (
        base / file_name,
        Path(__file__).resolve().parent / file_name,
    )
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


def logo_path() -> Path | None:
    path = resource_path(LOGO_FILE_NAME)
    return path if path.exists() else None


def normalize_header(value: object) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    return re.sub(r"[\s:：/\\()（）_-]+", "", text).strip()


def header_matches(value: object, expected_header: str) -> bool:
    text = normalize_header(value)
    if not text:
        return False
    aliases = HEADER_ALIASES.get(expected_header, (expected_header,))
    for alias in aliases:
        normalized_alias = normalize_header(alias)
        if text == normalized_alias or text.startswith(normalized_alias):
            return True
    return False


def locate_header_columns(row: tuple[object, ...], headers: tuple[str, ...]) -> dict[str, int]:
    columns: dict[str, int] = {}
    for column_index, value in enumerate(row, start=1):
        for header in headers:
            if header not in columns and header_matches(value, header):
                columns[header] = column_index
    return columns


def normalize_container(value: object) -> str:
    if value is None:
        return ""
    text = str(value).upper().strip()
    compact = re.sub(r"[^A-Z0-9]", "", text)
    match = CONTAINER_RE.search(compact)
    if match:
        return match.group(1)
    return compact if re.fullmatch(r"[A-Z]{4}\d{7}", compact) else ""


def container_check_digit(container: str) -> int | None:
    compact = re.sub(r"[^A-Z0-9]", "", container.upper())
    if not re.fullmatch(r"[A-Z]{4}\d{7}", compact):
        return None
    total = 0
    for index, character in enumerate(compact[:10]):
        value = CONTAINER_CHECK_VALUES.get(character)
        if value is None:
            return None
        total += value * (2**index)
    return (total % 11) % 10


def is_valid_container_number(container: str) -> bool:
    compact = re.sub(r"[^A-Z0-9]", "", container.upper())
    check_digit = container_check_digit(compact)
    return check_digit is not None and str(check_digit) == compact[-1]


def edit_distance(left: str, right: str) -> int:
    previous = list(range(len(right) + 1))
    for left_index, left_character in enumerate(left, start=1):
        current = [left_index] + [0] * len(right)
        for right_index, right_character in enumerate(right, start=1):
            current[right_index] = min(
                previous[right_index] + 1,
                current[right_index - 1] + 1,
                previous[right_index - 1] + (0 if left_character == right_character else 1),
            )
        previous = current
    return previous[-1]


def match_container_row(
    scanned_container: str,
    rows_by_container: dict[str, DriverExcelRow],
    rows: list[DriverExcelRow],
) -> RowMatch:
    scanned_container = normalize_container(scanned_container) or re.sub(r"[^A-Z0-9]", "", scanned_container.upper())
    if scanned_container in rows_by_container:
        return RowMatch(rows_by_container[scanned_container], "container_ocr", score=1.0, note="exact container OCR match")
    if len(scanned_container) != 11:
        return RowMatch(None, "", note=f"container OCR found '{scanned_container}', expected 11 characters")

    close_matches: list[tuple[int, DriverExcelRow]] = []
    for row in rows:
        distance = edit_distance(scanned_container, row.container)
        if distance <= 1 and is_valid_container_number(row.container):
            close_matches.append((distance, row))

    if len(close_matches) == 1:
        distance, row = close_matches[0]
        note = (
            f"Container OCR read '{scanned_container}', accepted '{row.container}' "
            f"because it differs by {distance} character and is the only close Excel container."
        )
        return RowMatch(row, "container_ocr_close", score=1.0 - (distance * 0.1), note=note)

    if close_matches:
        options = ", ".join(row.container for _distance, row in close_matches[:4])
        return RowMatch(None, "", note=f"container OCR close match is ambiguous: {options}")
    return RowMatch(None, "", note=f"container OCR read '{scanned_container}', no safe container match")


def parse_sequence(value: object) -> int | None:
    if value is None:
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    text = str(value).strip()
    if not text:
        return None
    match = re.search(r"\d+", text)
    return int(match.group(0)) if match else None


def load_waybill_rows(
    excel_path: Path,
    header_row_override: int | None = None,
    sequence_column_override: int | None = None,
    container_column_override: int | None = None,
) -> list[WaybillRow]:
    import_dependencies_ok()

    workbook = load_workbook(excel_path, data_only=True, read_only=True)
    worksheet = workbook.active

    manual_columns = sequence_column_override is not None or container_column_override is not None
    try:
        if manual_columns:
            if header_row_override is None or header_row_override < 1:
                raise ValueError("Manual Excel column setup needs a valid header row number.")
            if sequence_column_override is None or sequence_column_override < 1:
                raise ValueError("Manual Excel column setup needs a valid sequence column.")
            if container_column_override is None or container_column_override < 1:
                raise ValueError("Manual Excel column setup needs a valid container column.")
            header_row = header_row_override
            seq_col = sequence_column_override - 1
            container_col = container_column_override - 1
        else:
            header_row, columns = find_header_columns(worksheet, (SEQ_HEADER, CONTAINER_HEADER))
            seq_col = columns[SEQ_HEADER] - 1
            container_col = columns[CONTAINER_HEADER] - 1
    except Exception:
        workbook.close()
        raise

    rows: list[WaybillRow] = []
    seen_containers: dict[str, tuple[int, int]] = {}
    duplicate_containers: list[tuple[str, int, int]] = []
    missing_sequence_rows: list[tuple[int, str, str]] = []
    missing_container_rows: list[tuple[int, int]] = []
    sequence_occurrences: dict[int, int] = {}
    last_sequence: int | None = None
    last_sequence_row: int | None = None

    for excel_row, row in enumerate(
        worksheet.iter_rows(min_row=header_row + 1, values_only=True),
        start=header_row + 1,
    ):
        seq_value = row[seq_col] if seq_col < len(row) else None
        container_value = row[container_col] if container_col < len(row) else None
        sequence = parse_sequence(seq_value)
        container = normalize_container(container_value)
        if sequence is None and not container:
            continue
        if sequence is None:
            if last_sequence is None:
                missing_sequence_rows.append((excel_row, container, "no previous sequence to inherit"))
                continue
            sequence = last_sequence
            sequence_note = f"inherited sequence {sequence} from Excel row {last_sequence_row}"
        else:
            last_sequence = sequence
            last_sequence_row = excel_row
            sequence_note = ""
        if not container:
            missing_container_rows.append((excel_row, sequence))
            continue

        if container in seen_containers:
            first_row, first_sequence = seen_containers[container]
            if sequence == first_sequence:
                continue
            duplicate_containers.append((container, first_row, excel_row))
            continue

        seen_containers[container] = (excel_row, sequence)
        sequence_occurrences[sequence] = sequence_occurrences.get(sequence, 0) + 1
        rows.append(
            WaybillRow(
                sequence=sequence,
                container=container,
                excel_row=excel_row,
                source_order=len(rows) + 1,
                sequence_occurrence=sequence_occurrences[sequence],
                sequence_note=sequence_note,
            )
        )

    validation_messages: list[str] = []
    if duplicate_containers:
        details = "; ".join(
            f"{container} in rows {first_row} and {second_row}"
            for container, first_row, second_row in duplicate_containers
        )
        validation_messages.append(f"Duplicate Excel container numbers: {details}.")
    if missing_sequence_rows:
        details = "; ".join(
            f"row {row} has container {container} ({reason})"
            for row, container, reason in missing_sequence_rows
        )
        validation_messages.append(f"Container rows missing sequence numbers: {details}.")
    if missing_container_rows:
        details = "; ".join(f"row {row} has sequence {sequence}" for row, sequence in missing_container_rows)
        validation_messages.append(f"Sequence rows missing container numbers: {details}.")
    if validation_messages:
        workbook.close()
        raise ValueError("\n".join(validation_messages))

    if not rows:
        workbook.close()
        raise ValueError("No rows with both sequence and container number were found.")

    rows.sort(key=lambda item: (item.sequence, item.source_order))
    workbook.close()
    return rows


def find_header_columns(worksheet: object, required_headers: tuple[str, ...]) -> tuple[int, dict[str, int]]:
    best_found: dict[str, int] = {}
    for row_index, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
        found = locate_header_columns(tuple(row), required_headers)
        if len(found) > len(best_found):
            best_found = found
        if all(header in found for header in required_headers):
            return row_index, found
    missing = ", ".join(header for header in required_headers if header not in best_found)
    raise ValueError(f"Could not find required Excel columns: {missing}")


def find_header_column_on_row(worksheet: object, header_row: int, header: str) -> int | None:
    for row in worksheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True):
        return locate_header_columns(tuple(row), (header,)).get(header)
    return None


def excel_column_name(index: int) -> str:
    if index < 1:
        return ""
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def parse_excel_column(value: object) -> int | None:
    text = str(value or "").strip().upper()
    if not text:
        return None
    if text.isdigit():
        number = int(text)
        return number if number > 0 else None
    if not re.fullmatch(r"[A-Z]+", text):
        return None
    number = 0
    for character in text:
        number = number * 26 + (ord(character) - 64)
    return number


def load_driver_excel_rows(excel_path: Path) -> tuple[list[DriverExcelRow], int, int, int]:
    import_dependencies_ok()

    workbook = load_workbook(excel_path, data_only=True, read_only=True)
    worksheet = workbook.active
    header_row, columns = find_header_columns(worksheet, (SEQ_HEADER, CONTAINER_HEADER))
    seq_col = columns[SEQ_HEADER]
    container_col = columns[CONTAINER_HEADER]
    car_col = find_header_column_on_row(worksheet, header_row, CAR_NUMBER_HEADER)

    rows: list[DriverExcelRow] = []
    blank_count = 0
    filled_count = 0
    for excel_row in range(header_row + 1, worksheet.max_row + 1):
        sequence = parse_sequence(worksheet.cell(excel_row, seq_col).value)
        container = normalize_container(worksheet.cell(excel_row, container_col).value)
        if sequence is None and not container:
            continue
        if sequence is None or not container:
            continue
        current_value = worksheet.cell(excel_row, car_col).value if car_col is not None else None
        current_text = "" if current_value is None else str(current_value).strip()
        if current_text:
            filled_count += 1
        else:
            blank_count += 1
        rows.append(
            DriverExcelRow(
                sequence=sequence,
                excel_row=excel_row,
                container=container,
                current_car_number=current_text,
            )
        )

    rows.sort(key=lambda item: item.sequence)
    if not rows:
        raise ValueError("No Excel rows with both sequence and container number were found.")
    return rows, header_row, blank_count, filled_count


def list_scan_images(image_folder: Path) -> list[Path]:
    return sorted(
        (
            item
            for item in Path(image_folder).iterdir()
            if item.is_file() and item.suffix.lower() in IMAGE_EXTENSIONS
        ),
        key=lambda path: path.name.lower(),
    )


OCR_TEMPLATE_SIZE = (36, 56)
DRIVER_NUMBER_REGION = (700, 880, 1350, 1160)
CONTAINER_SCAN_REGION = (70, 1800, 900, 2350)
OCR_TEMPLATE_CACHE: dict[str, dict[str, list[frozenset[int]]]] = {}


def ocr_dependencies_ok() -> None:
    if Image is None or ImageDraw is None or ImageFont is None:
        raise RuntimeError("Image reading is unavailable. Run the packaged .exe or install Pillow.")


def template_font_paths() -> tuple[Path, ...]:
    return (
        Path("C:/Windows/Fonts/arial.ttf"),
        Path("C:/Windows/Fonts/arialbd.ttf"),
        Path("C:/Windows/Fonts/calibri.ttf"),
        Path("C:/Windows/Fonts/calibrib.ttf"),
    )


def image_black_indices(image: object, threshold: int = 150) -> frozenset[int]:
    gray = image.convert("L")
    width, _height = gray.size
    return frozenset(index for index, value in enumerate(gray.getdata()) if value < threshold)


def normalize_glyph_image(image: object, threshold: int = 145) -> frozenset[int]:
    gray = image.convert("L")
    width, height = gray.size
    black = [(index % width, index // width) for index, value in enumerate(gray.getdata()) if value < threshold]
    if not black:
        return frozenset()
    xs = [point[0] for point in black]
    ys = [point[1] for point in black]
    glyph = gray.crop((min(xs), min(ys), max(xs) + 1, max(ys) + 1))
    canvas = Image.new("L", OCR_TEMPLATE_SIZE, 255)
    glyph.thumbnail((OCR_TEMPLATE_SIZE[0] - 4, OCR_TEMPLATE_SIZE[1] - 4), Image.Resampling.LANCZOS)
    canvas.paste(glyph, ((OCR_TEMPLATE_SIZE[0] - glyph.width) // 2, (OCR_TEMPLATE_SIZE[1] - glyph.height) // 2))
    return image_black_indices(canvas, 150)


def build_ocr_templates(characters: str) -> dict[str, list[frozenset[int]]]:
    ocr_dependencies_ok()
    cache_key = "".join(sorted(set(characters)))
    if cache_key in OCR_TEMPLATE_CACHE:
        return OCR_TEMPLATE_CACHE[cache_key]

    templates: dict[str, list[frozenset[int]]] = {character: [] for character in cache_key}
    for font_path in template_font_paths():
        if not font_path.exists():
            continue
        for size in range(34, 76, 3):
            try:
                font = ImageFont.truetype(str(font_path), size)
            except Exception:
                continue
            for character in cache_key:
                image = Image.new("L", (100, 100), 255)
                draw = ImageDraw.Draw(image)
                bbox = draw.textbbox((0, 0), character, font=font)
                draw.text((-bbox[0] + 5, -bbox[1] + 5), character, font=font, fill=0)
                glyph = normalize_glyph_image(image)
                if glyph:
                    templates[character].append(glyph)

    OCR_TEMPLATE_CACHE[cache_key] = templates
    return templates


def connected_components(mask: bytearray, width: int, height: int) -> list[tuple[int, int, int, int, int]]:
    seen = bytearray(width * height)
    components: list[tuple[int, int, int, int, int]] = []
    for start, value in enumerate(mask):
        if not value or seen[start]:
            continue
        stack = [start]
        seen[start] = 1
        min_x = max_x = start % width
        min_y = max_y = start // width
        area = 0
        while stack:
            index = stack.pop()
            area += 1
            x = index % width
            y = index // width
            min_x = min(min_x, x)
            max_x = max(max_x, x)
            min_y = min(min_y, y)
            max_y = max(max_y, y)
            for ny in (y - 1, y, y + 1):
                if ny < 0 or ny >= height:
                    continue
                row = ny * width
                for nx in (x - 1, x, x + 1):
                    if nx < 0 or nx >= width:
                        continue
                    neighbor = row + nx
                    if not seen[neighbor] and mask[neighbor]:
                        seen[neighbor] = 1
                        stack.append(neighbor)
        components.append((min_x, min_y, max_x + 1, max_y + 1, area))
    return components


def threshold_mask(image: object, threshold: int) -> bytearray:
    gray = image.convert("L")
    return bytearray(1 if value < threshold else 0 for value in gray.getdata())


def recognize_glyph(mask: frozenset[int], allowed: str) -> tuple[str, float]:
    templates = build_ocr_templates(allowed)
    best_character = ""
    best_score = 0.0
    for character in allowed:
        for template in templates.get(character, []):
            intersection = len(mask & template)
            union = len(mask) + len(template) - intersection
            score = intersection / union if union else 0.0
            if score > best_score:
                best_character = character
                best_score = score
    return best_character, best_score


def region_black_density(image: object, x0: float, x1: float, y0: float, y1: float, threshold: int = 130) -> float:
    gray = image.convert("L")
    width, height = gray.size
    start_x = max(0, min(width - 1, int(width * x0)))
    end_x = max(start_x + 1, min(width, int(width * x1)))
    start_y = max(0, min(height - 1, int(height * y0)))
    end_y = max(start_y + 1, min(height, int(height * y1)))
    pixels = gray.load()
    total = 0
    black = 0
    for y in range(start_y, end_y):
        for x in range(start_x, end_x):
            total += 1
            if pixels[x, y] < threshold:
                black += 1
    return black / total if total else 0.0


def correct_digit_shape(image: object, character: str) -> str:
    if character != "8":
        return character
    upper_left = region_black_density(image, 0.0, 0.35, 0.10, 0.45)
    upper_right = region_black_density(image, 0.65, 1.0, 0.10, 0.45)
    lower_left = region_black_density(image, 0.0, 0.35, 0.55, 0.90)
    lower_right = region_black_density(image, 0.65, 1.0, 0.55, 0.90)
    middle = region_black_density(image, 0.20, 0.80, 0.40, 0.62)
    if upper_left > 0.45 and upper_right < 0.25 and lower_left < 0.22 and lower_right > 0.40 and middle > 0.35:
        return "5"
    return character


def split_wide_digit_components(components: list[tuple[int, int, int, int, int]]) -> list[tuple[int, int, int, int, int]]:
    if not components:
        return []
    widths = sorted(component[2] - component[0] for component in components if component[2] - component[0] <= 60)
    median_width = widths[len(widths) // 2] if widths else 36
    split: list[tuple[int, int, int, int, int]] = []
    for x0, y0, x1, y1, area in sorted(components, key=lambda item: item[0]):
        width = x1 - x0
        parts = max(1, round(width / max(1, median_width)))
        if parts > 1 and width > median_width * 1.55:
            for part in range(parts):
                part_x0 = x0 + round(width * part / parts)
                part_x1 = x0 + round(width * (part + 1) / parts)
                split.append((part_x0, y0, part_x1, y1, max(1, area // parts)))
        else:
            split.append((x0, y0, x1, y1, area))
    return split


WINDOWS_OCR_SCRIPT = r"""
param(
    [Parameter(Mandatory = $true)][string]$InputJson,
    [Parameter(Mandatory = $true)][string]$OutputJson
)

$ErrorActionPreference = "Stop"
$paths = @(Get-Content -Encoding UTF8 $InputJson)
$results = New-Object System.Collections.Generic.List[object]

try {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
    $null = [Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
    $null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType = WindowsRuntime]

    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object { $_.Name -eq "AsTask" -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 })[0]

    function AwaitOperation($async, [Type]$resultType) {
        $asTask = $asTaskGeneric.MakeGenericMethod($resultType)
        $task = $asTask.Invoke($null, @($async))
        $task.Wait() | Out-Null
        $task.Result
    }

    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if ($null -eq $engine) {
        throw "Windows OCR engine is not available for the current user profile languages."
    }

    foreach ($pathItem in $paths) {
        $path = [string]$pathItem
        try {
            $file = AwaitOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($path)) ([Windows.Storage.StorageFile])
            $stream = AwaitOperation ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
            try {
                $decoder = AwaitOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
                $bitmap = AwaitOperation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
                $result = AwaitOperation ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
                $results.Add([pscustomobject]@{ path = $path; text = $result.Text; error = "" })
            }
            finally {
                if ($stream -and ($stream -is [System.IDisposable])) {
                    $stream.Dispose()
                }
            }
        }
        catch {
            $results.Add([pscustomobject]@{ path = $path; text = ""; error = $_.Exception.Message })
        }
    }
}
catch {
    foreach ($pathItem in $paths) {
        $path = [string]$pathItem
        $results.Add([pscustomobject]@{ path = $path; text = ""; error = $_.Exception.Message })
    }
}

$results | ConvertTo-Json -Depth 4 | Set-Content -Encoding UTF8 $OutputJson
"""


def read_windows_ocr_texts(image_paths: list[Path], log: Callable[[str], None] | None = None) -> dict[str, str]:
    if not image_paths:
        return {}

    with tempfile.TemporaryDirectory(prefix="unitrans_ocr_") as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        input_path = temp_dir / "images.json"
        output_path = temp_dir / "ocr.json"
        script_path = temp_dir / "ocr.ps1"
        input_path.write_text("\n".join(str(path) for path in image_paths), encoding="utf-8")
        script_path.write_text(WINDOWS_OCR_SCRIPT, encoding="utf-8")

        command = [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_path),
            str(input_path),
            str(output_path),
        ]
        startup_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        timeout_seconds = max(60, len(image_paths) * 4)
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=timeout_seconds,
                creationflags=startup_flags,
            )
        except Exception as exc:
            if log:
                log(f"Windows OCR unavailable; using built-in image OCR only. {exc}")
            return {}

        if completed.returncode != 0 or not output_path.exists():
            detail = completed.stderr.strip() or completed.stdout.strip() or f"exit code {completed.returncode}"
            if log:
                log(f"Windows OCR unavailable; using built-in image OCR only. {detail}")
            return {}

        try:
            payload = json.loads(output_path.read_text(encoding="utf-8-sig"))
        except Exception as exc:
            if log:
                log(f"Windows OCR result could not be read; using built-in image OCR only. {exc}")
            return {}

    if isinstance(payload, dict):
        payload = [payload]

    texts: dict[str, str] = {}
    error_count = 0
    for item in payload:
        if not isinstance(item, dict):
            continue
        path = str(item.get("path") or "")
        text = str(item.get("text") or "")
        error = str(item.get("error") or "")
        if path and text:
            texts[path] = text
        elif error:
            error_count += 1

    if log:
        log(f"Windows OCR completed: {len(texts)}/{len(image_paths)} image(s) returned text.")
        if error_count:
            log(f"Windows OCR skipped {error_count} image(s); built-in image OCR will handle them.")
    return texts


def container_text_variants(container: str) -> set[str]:
    container = normalize_container(container)
    if not container:
        return set()

    replacements = {
        "I": ("I", "L", "1"),
        "O": ("O", "0"),
        "S": ("S", "5"),
        "B": ("B", "8"),
        "Z": ("Z", "2"),
    }
    variants = [""]
    for index, character in enumerate(container):
        choices = replacements.get(character, (character,)) if index < 4 else (character,)
        variants = [prefix + choice for prefix in variants for choice in choices]
    return set(variants)


def expected_containers_in_ocr_text(text: str, expected_containers: set[str]) -> list[str]:
    if not text or not expected_containers:
        return []

    compact = re.sub(r"[^A-Z0-9]", "", text.upper())
    matches = []
    for container in sorted(expected_containers):
        if any(variant in compact for variant in container_text_variants(container)):
            matches.append(container)
    return matches


def extract_driver_number_from_image(image_path: Path) -> OcrTextResult:
    ocr_dependencies_ok()
    image = Image.open(image_path)
    crop = image.crop(DRIVER_NUMBER_REGION)
    width, height = crop.size
    mask = threshold_mask(crop, 100)
    likely: list[tuple[int, int, int, int, int]] = []
    for component in connected_components(mask, width, height):
        x0, y0, x1, y1, area = component
        component_width = x1 - x0
        component_height = y1 - y0
        if component_height > 45 and component_width > 15 and area > 300 and y0 > 70:
            likely.append(component)

    likely = split_wide_digit_components(likely)
    if len(likely) > 7:
        likely = sorted(sorted(likely, key=lambda item: item[4], reverse=True)[:7], key=lambda item: item[0])

    digits: list[str] = []
    scores: list[float] = []
    for component in sorted(likely, key=lambda item: item[0]):
        glyph_image = crop.crop(component[:4])
        glyph = normalize_glyph_image(glyph_image)
        character, score = recognize_glyph(glyph, "0123456789")
        if character:
            digits.append(correct_digit_shape(glyph_image, character))
            scores.append(score)

    text = "".join(digits)
    confidence = min(scores) if scores else 0.0
    note = "" if len(text) == 7 else f"expected 7 digits, found {len(text)}"
    return OcrTextResult(text=text, confidence=confidence, note=note)


def cluster_container_line(
    components: list[tuple[int, int, int, int, int]]
) -> list[tuple[int, int, int, int, int]]:
    clusters: list[list[tuple[int, int, int, int, int]]] = []
    for component in sorted(components, key=lambda item: (item[1] + item[3]) / 2):
        center_y = (component[1] + component[3]) / 2
        for cluster in clusters:
            cluster_center = sum((item[1] + item[3]) / 2 for item in cluster) / len(cluster)
            if abs(center_y - cluster_center) < 22:
                cluster.append(component)
                break
        else:
            clusters.append([component])

    candidates: list[tuple[float, float, list[tuple[int, int, int, int, int]]]] = []
    for cluster in clusters:
        left_side = sorted((item for item in cluster if item[0] < 650), key=lambda item: item[0])
        if len(left_side) < 11:
            continue
        heights = sorted(item[3] - item[1] for item in left_side)
        median_height = heights[len(heights) // 2]
        center_y = sum((item[1] + item[3]) / 2 for item in left_side) / len(left_side)
        candidates.append((median_height, -center_y, left_side))
    if not candidates:
        return []
    candidates.sort(reverse=True)
    return sorted(candidates[0][2], key=lambda item: item[0])[:11]


def extract_container_from_scan_image_by_template(image_path: Path, note_prefix: str = "") -> OcrTextResult:
    ocr_dependencies_ok()
    image = Image.open(image_path)
    crop = image.crop(CONTAINER_SCAN_REGION)
    width, height = crop.size
    mask = threshold_mask(crop, 130)
    candidates: list[tuple[int, int, int, int, int]] = []
    for component in connected_components(mask, width, height):
        x0, y0, x1, y1, area = component
        component_width = x1 - x0
        component_height = y1 - y0
        if 4 <= component_width <= 90 and 22 <= component_height <= 70 and area >= 35:
            candidates.append(component)

    line = cluster_container_line(candidates)
    characters: list[str] = []
    scores: list[float] = []
    for index, component in enumerate(line):
        allowed = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" if index < 4 else "0123456789"
        glyph = normalize_glyph_image(crop.crop(component[:4]))
        character, score = recognize_glyph(glyph, allowed)
        if character:
            characters.append(character)
            scores.append(score)

    text = "".join(characters)
    confidence = min(scores) if scores else 0.0
    note = "" if len(text) == 11 else f"expected 11 characters, found {len(text)}"
    if note_prefix:
        note = "; ".join(item for item in (note_prefix, note, "built-in image OCR") if item)
    return OcrTextResult(text=text, confidence=confidence, note=note)


def extract_container_from_scan_image(
    image_path: Path,
    expected_containers: set[str] | None = None,
    windows_ocr_text: str | None = None,
) -> OcrTextResult:
    expected_containers = expected_containers or set()
    if windows_ocr_text and expected_containers:
        matches = expected_containers_in_ocr_text(windows_ocr_text, expected_containers)
        if len(matches) == 1:
            return OcrTextResult(text=matches[0], confidence=1.0, note="windows_ocr")
        if len(matches) > 1:
            note = "Windows OCR found multiple expected containers: " + ", ".join(matches[:6])
            return extract_container_from_scan_image_by_template(image_path, note)
        return extract_container_from_scan_image_by_template(image_path, "Windows OCR found no expected container")

    return extract_container_from_scan_image_by_template(image_path)


def is_blank_cell(value: object) -> bool:
    return value is None or str(value).strip() == ""


def fill_driver_numbers_from_images(
    excel_path: Path,
    image_folder: Path,
    output_root: Path | None = None,
    log: Callable[[str], None] | None = None,
) -> DriverFillResult:
    import_dependencies_ok()
    ocr_dependencies_ok()

    def emit(message: str) -> None:
        if log:
            try:
                log(message)
            except UnicodeEncodeError:
                try:
                    log(message.encode("ascii", errors="replace").decode("ascii"))
                except Exception:
                    pass

    excel_path = Path(excel_path)
    image_folder = Path(image_folder)
    if not excel_path.is_file():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    if not image_folder.is_dir():
        raise FileNotFoundError(f"Image folder not found: {image_folder}")

    driver_rows, header_row, blank_count, _already_filled_in_source = load_driver_excel_rows(excel_path)
    image_paths = list_scan_images(image_folder)
    if not image_paths:
        raise ValueError("No image files were found in the selected folder.")

    emit(f"Driver workflow Excel rows: {len(driver_rows)}; blank car numbers: {blank_count}.")
    emit(f"Scanning {len(image_paths)} waybill images.")

    workbook = load_workbook(excel_path)
    worksheet = workbook.active
    write_header_row, columns = find_header_columns(worksheet, (SEQ_HEADER, CONTAINER_HEADER))
    car_col = find_header_column_on_row(worksheet, write_header_row, CAR_NUMBER_HEADER)
    if car_col is None:
        car_col = max(worksheet.max_column + 1, max(columns.values()) + 1)
        worksheet.cell(write_header_row, car_col).value = CAR_NUMBER_HEADER
        emit("Created missing car number column: 车号.")

    rows_by_container = {row.container: row for row in driver_rows}
    expected_container_set = set(rows_by_container)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = Path(output_root) if output_root is not None else image_folder / f"driver_numbers_{timestamp}"
    output_dir.mkdir(parents=True, exist_ok=False)
    workbook_path = unique_path(output_dir / f"{excel_path.stem}_driver_numbers_filled.xlsx")
    report_path = output_dir / "driver_number_report.csv"

    report_rows: list[list[str]] = []
    filled_count = 0
    already_filled_count = 0
    unmatched_count = 0
    review_count = 0
    windows_ocr_texts = read_windows_ocr_texts(image_paths, emit)
    scan_results: list[tuple[Path, OcrTextResult, OcrTextResult]] = []
    for image_path in image_paths:
        scan_results.append(
            (
                image_path,
                extract_driver_number_from_image(image_path),
                extract_container_from_scan_image(
                    image_path,
                    expected_container_set,
                    windows_ocr_texts.get(str(image_path)),
                ),
            )
        )

    exact_container_matches = sum(1 for _path, _driver, scanned in scan_results if scanned.text in rows_by_container)
    windows_container_matches = sum(
        1
        for _path, _driver, scanned in scan_results
        if scanned.note.startswith("windows_ocr") and scanned.text in rows_by_container
    )
    close_container_matches = 0

    for index, (image_path, driver, scanned_container) in enumerate(scan_results):
        container_match = match_container_row(scanned_container.text, rows_by_container, driver_rows)
        target_row = container_match.row
        match_method = container_match.method
        review_note = container_match.note
        if target_row is not None and match_method == "container_ocr" and scanned_container.note.startswith("windows_ocr"):
            match_method = "windows_ocr"
            review_note = "Windows OCR matched an expected Excel container."
        if match_method == "container_ocr_close":
            close_container_matches += 1

        if target_row is None:
            unmatched_count += 1
            review_count += 1
            report_rows.append(
                [
                    image_path.name,
                    "",
                    scanned_container.text,
                    "",
                    driver.text,
                    f"{driver.confidence:.3f}",
                    "review_required",
                    review_note or "no matching Excel row",
                ]
            )
            emit(f"[{index + 1}/{len(image_paths)}] {image_path.name} -> review required")
            continue

        if not re.fullmatch(r"\d{7}", driver.text):
            unmatched_count += 1
            report_rows.append(
                [
                    image_path.name,
                    str(target_row.sequence),
                    scanned_container.text,
                    target_row.container,
                    driver.text,
                    f"{driver.confidence:.3f}",
                    match_method,
                    "; ".join(item for item in (driver.note, review_note, "driver number not filled") if item),
                ]
            )
            emit(f"{image_path.name}: driver number could not be read confidently.")
            continue

        cell = worksheet.cell(target_row.excel_row, car_col)
        if is_blank_cell(cell.value):
            cell.value = driver.text
            filled_count += 1
            action = "filled"
        else:
            already_filled_count += 1
            action = "already filled"

        notes = "; ".join(item for item in (driver.note, scanned_container.note, review_note, action) if item)
        report_rows.append(
            [
                image_path.name,
                str(target_row.sequence),
                scanned_container.text,
                target_row.container,
                driver.text,
                f"{driver.confidence:.3f}",
                match_method,
                notes,
            ]
        )
        emit(f"[{index + 1}/{len(image_paths)}] {image_path.name} -> {target_row.container}: {driver.text} ({match_method})")

    workbook.save(workbook_path)
    with report_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "image_file",
                "sequence",
                "scanned_container",
                "excel_container",
                "driver_number",
                "driver_confidence",
                "match_method",
                "notes",
            ]
        )
        writer.writerows(report_rows)

    emit(f"Driver numbers filled: {filled_count}. Already filled: {already_filled_count}.")
    emit(f"Windows OCR container matches: {windows_container_matches}.")
    emit(f"Container OCR exact matches: {exact_container_matches}; close container matches: {close_container_matches}.")
    if review_count:
        emit(f"Review required: {review_count} image(s) were not filled automatically.")
    emit(f"Driver workbook: {workbook_path}")

    return DriverFillResult(
        output_dir=output_dir,
        workbook_path=workbook_path,
        report_path=report_path,
        expected_count=len(driver_rows),
        image_count=len(image_paths),
        filled_count=filled_count,
        already_filled_count=already_filled_count,
        unmatched_count=unmatched_count,
        review_count=review_count,
    )


def extract_container_candidates(text: str) -> set[str]:
    compact = re.sub(r"[^A-Z0-9]", "", text.upper())
    return set(CONTAINER_RE.findall(compact))


def extract_pdf_text(pdf_path: Path, max_pages: int | None = None) -> str:
    import_dependencies_ok()

    reader = PdfReader(str(pdf_path))
    if reader.is_encrypted:
        decrypt_result = reader.decrypt("")
        if decrypt_result == 0:
            raise ValueError("PDF is encrypted and could not be opened with an empty password.")
    page_count = len(reader.pages)
    limit = page_count if max_pages is None else min(page_count, max_pages)
    return "\n".join(reader.pages[index].extract_text() or "" for index in range(limit))


def identify_pdf(pdf_path: Path, expected_containers: set[str]) -> PdfMatch:
    notes: list[str] = []
    prefix_match = re.match(r"^([A-Z]{4}\d{7})(?:[^A-Z0-9]|$)", pdf_path.name.upper())
    if prefix_match and prefix_match.group(1) in expected_containers:
        return PdfMatch(pdf_path, prefix_match.group(1), "filename", "")

    filename_matches = sorted(extract_container_candidates(pdf_path.stem) & expected_containers)
    if len(filename_matches) == 1:
        return PdfMatch(pdf_path, filename_matches[0], "filename", "")
    if len(filename_matches) > 1:
        return PdfMatch(
            pdf_path,
            None,
            "ambiguous",
            "matched multiple containers in filename: " + ", ".join(filename_matches),
        )

    try:
        text = extract_pdf_text(pdf_path, max_pages=1)
    except Exception as exc:
        text = ""
        notes.append(f"text extraction failed: {exc}")

    matches: list[str] = []
    if text:
        candidates = extract_container_candidates(text)
        matches = sorted(candidates & expected_containers)
        if not matches:
            compact_text = re.sub(r"[^A-Z0-9]", "", text.upper())
            matches = sorted(container for container in expected_containers if container in compact_text)

    if len(matches) == 1:
        return PdfMatch(pdf_path, matches[0], "text page 1", "; ".join(notes))
    if len(matches) > 1:
        return PdfMatch(pdf_path, None, "ambiguous", "matched multiple containers in PDF text: " + ", ".join(matches))

    if text:
        try:
            full_text = extract_pdf_text(pdf_path)
        except Exception as exc:
            notes.append(f"full text fallback failed: {exc}")
        else:
            full_matches = sorted(extract_container_candidates(full_text) & expected_containers)
            if len(full_matches) == 1:
                return PdfMatch(pdf_path, full_matches[0], "full text", "; ".join(notes))
            if len(full_matches) > 1:
                return PdfMatch(
                    pdf_path,
                    None,
                    "ambiguous",
                    "matched multiple containers in full PDF text: " + ", ".join(full_matches),
                )

    note = "; ".join(notes + ["no expected container found"])
    return PdfMatch(pdf_path, None, "unmatched", note)


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    for index in range(2, 1000):
        candidate = path.with_name(f"{stem}_{index}{suffix}")
        if not candidate.exists():
            return candidate
    raise FileExistsError(f"Too many duplicate output files for {path.name}.")


def write_summary(path: Path, lines: Iterable[str]) -> None:
    path.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_decrypted_pdf(reader: object, output_path: Path) -> None:
    writer = PdfWriter()
    try:
        writer.clone_document_from_reader(reader)
    except Exception:
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        try:
            if reader.metadata:
                writer.add_metadata({key: str(value) for key, value in reader.metadata.items() if value is not None})
        except Exception:
            pass
    with output_path.open("wb") as handle:
        writer.write(handle)


UNLOCKABLE_EXTENSIONS = {".pdf", ".xlsx", ".xlsm", ".xltx", ".xltm"}


def _unlock_spreadsheet(source: Path, destination: Path, password: str) -> tuple[bool, bool, str]:
    """Return (changed, skipped, note) for an Excel Open XML workbook."""
    encrypted = False
    decrypted_path: Path | None = None
    with source.open("rb") as source_handle:
        office_file = msoffcrypto.OfficeFile(source_handle)
        encrypted = office_file.is_encrypted()
        if encrypted:
            if not password:
                return False, True, "Workbook requires an open password; enter it and retry."
            try:
                office_file.load_key(password=password)
                with tempfile.NamedTemporaryFile(
                    prefix="unitrans-decrypted-", suffix=source.suffix,
                    dir=destination.parent, delete=False,
                ) as output_handle:
                    decrypted_path = Path(output_handle.name)
                    office_file.decrypt(output_handle)
            except Exception:
                if decrypted_path:
                    decrypted_path.unlink(missing_ok=True)
                return False, True, "The workbook password was not accepted; source was not changed."

    removed_protection = False
    try:
        input_path = decrypted_path or source
        with zipfile.ZipFile(input_path, "r") as source_zip, zipfile.ZipFile(destination, "w") as output_zip:
            for entry in source_zip.infolist():
                content = source_zip.read(entry.filename)
                if entry.filename == "xl/workbook.xml":
                    content, count = re.subn(rb"<(?:[A-Za-z0-9_]+:)?workbookProtection\b[^>]*/\s*>", b"", content)
                    removed_protection = removed_protection or count > 0
                elif entry.filename.startswith("xl/worksheets/") and entry.filename.endswith(".xml"):
                    content, count = re.subn(rb"<(?:[A-Za-z0-9_]+:)?sheetProtection\b[^>]*/\s*>", b"", content)
                    removed_protection = removed_protection or count > 0
                output_zip.writestr(entry, content)
        with zipfile.ZipFile(destination, "r") as verify_zip:
            bad_entry = verify_zip.testzip()
            if bad_entry:
                raise ValueError(f"Workbook archive verification failed at {bad_entry}.")
    except Exception:
        destination.unlink(missing_ok=True)
        raise
    finally:
        if decrypted_path:
            decrypted_path.unlink(missing_ok=True)

    if encrypted and removed_protection:
        return True, False, "Removed file-open encryption and workbook/sheet protection."
    if encrypted:
        return True, False, "Removed file-open encryption."
    if removed_protection:
        return True, False, "Removed workbook and/or worksheet protection."
    destination.unlink(missing_ok=True)
    shutil.copy2(source, destination)
    return False, False, "Copied unchanged; no workbook or worksheet protection was found."


def unlock_files(
    sources: Iterable[Path],
    password: str = "",
    output_root: Path | None = None,
    log: Callable[[str], None] | None = None,
    progress: Callable[[int, int], None] | None = None,
) -> PdfUnlockResult:
    import_dependencies_ok()

    def emit(message: str) -> None:
        if log:
            log(message)

    selected_paths = tuple(Path(source) for source in sources)
    source_paths: list[Path] = []
    for source in selected_paths:
        if source.is_dir():
            source_paths.extend(
                item for item in source.rglob("*")
                if item.is_file()
                and item.suffix.lower() in UNLOCKABLE_EXTENSIONS
                and not any(part.startswith("unlocked_files_") for part in item.relative_to(source).parts)
            )
        elif source.is_file() and source.suffix.lower() in UNLOCKABLE_EXTENSIONS:
            source_paths.append(source)
    source_paths = sorted(dict.fromkeys(path.resolve() for path in source_paths), key=lambda item: str(item).lower())
    if not source_paths:
        raise ValueError("No supported PDF or Excel files were found in the selection.")

    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    if output_root is not None:
        base_output = Path(output_root)
    elif len(selected_paths) == 1 and selected_paths[0].is_dir():
        base_output = selected_paths[0]
    else:
        base_output = source_paths[0].parent
    output_dir = unique_path(base_output / f"unlocked_files_{timestamp}")
    unlocked_dir = output_dir / "Unlocked Files"
    unlocked_dir.mkdir(parents=True, exist_ok=True)
    report_path = output_dir / "unlock_report.csv"
    summary_path = output_dir / "summary.txt"

    emit(f"Scanning {len(source_paths)} PDF and Excel files for removable protection.")

    report_rows: list[dict[str, object]] = []
    unlocked_count = 0
    already_unlocked_count = 0
    skipped_count = 0
    failed_count = 0

    used_output_paths: set[Path] = set()
    for index, source_path in enumerate(source_paths, start=1):
        output_path = unlocked_dir / source_path.name
        if output_path in used_output_paths:
            output_path = output_path.with_name(f"{source_path.stem}_{hashlib.sha1(str(source_path).encode()).hexdigest()[:8]}{source_path.suffix}")
        used_output_paths.add(output_path)
        status = ""
        note = ""
        try:
            if source_path.suffix.lower() == ".pdf":
                reader = PdfReader(str(source_path), strict=False)
                if reader.is_encrypted:
                    decrypt_result = reader.decrypt(password)
                    if not decrypt_result:
                        status = "skipped_password_required"
                        note = "PDF requires a valid open password and was not changed."
                        skipped_count += 1
                    else:
                        write_decrypted_pdf(reader, output_path)
                        verify_reader = PdfReader(str(output_path), strict=False)
                        if verify_reader.is_encrypted:
                            raise ValueError("Output PDF is still encrypted after rewrite.")
                        status = "unlocked"
                        note = "Removed PDF encryption."
                else:
                    shutil.copy2(source_path, output_path)
                    status = "already_unlocked"
                    note = "Copied unchanged; source PDF was not encrypted."
            else:
                changed, skipped, note = _unlock_spreadsheet(source_path, output_path, password)
                if skipped:
                    status = "skipped_password_required"
                    skipped_count += 1
                elif changed:
                    status = "unlocked"
                else:
                    status = "already_unlocked"

            if status == "unlocked":
                unlocked_count += 1
                emit(f"[{index}/{len(source_paths)}] {source_path.name} -> unlocked")
            elif status == "already_unlocked":
                already_unlocked_count += 1
                emit(f"[{index}/{len(source_paths)}] {source_path.name} -> already unlocked")
            else:
                emit(f"[{index}/{len(source_paths)}] {source_path.name} -> skipped; password required")
        except Exception as exc:
            status = "failed"
            note = str(exc)
            failed_count += 1
            emit(f"[{index}/{len(source_paths)}] {source_path.name} -> failed: {exc}")

        report_rows.append(
            {
                "source_file": str(source_path),
                "output_file": str(output_path) if status in {"unlocked", "already_unlocked"} else "",
                "status": status,
                "file_type": source_path.suffix.lstrip(".").upper(),
                "note": note,
            }
        )
        if progress:
            progress(index, len(source_paths))

    with report_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=("source_file", "output_file", "status", "file_type", "note"))
        writer.writeheader()
        writer.writerows(report_rows)

    write_summary(
        summary_path,
        (
            f"Selected files: {len(source_paths)}",
            f"Output folder: {output_dir}",
            f"PDF and Excel files scanned: {len(source_paths)}",
            f"Files unlocked: {unlocked_count}",
            f"Already unlocked files copied: {already_unlocked_count}",
            f"Skipped password-required files: {skipped_count}",
            f"Failed files: {failed_count}",
            f"Report: {report_path}",
        ),
    )

    emit(
        "File unlock finished: "
        f"{unlocked_count} unlocked; "
        f"{already_unlocked_count} already unlocked; "
        f"{skipped_count} skipped; "
        f"{failed_count} failed."
    )

    return PdfUnlockResult(
        output_dir=output_dir,
        unlocked_dir=unlocked_dir,
        report_path=report_path,
        summary_path=summary_path,
        total_count=len(source_paths),
        unlocked_count=unlocked_count,
        already_unlocked_count=already_unlocked_count,
        skipped_count=skipped_count,
        failed_count=failed_count,
    )


def unlock_pdf_folder(
    pdf_folder: Path,
    output_root: Path | None = None,
    log: Callable[[str], None] | None = None,
    progress: Callable[[int, int], None] | None = None,
) -> PdfUnlockResult:
    """Compatibility wrapper for the original PDF-folder workflow and CLI."""
    return unlock_files((Path(pdf_folder),), output_root=output_root, log=log, progress=progress)


def sort_waybills(
    excel_path: Path,
    pdf_folder: Path,
    output_root: Path | None = None,
    log: Callable[[str], None] | None = None,
) -> SortResult:
    def emit(message: str) -> None:
        if log:
            log(message)

    excel_path = Path(excel_path)
    pdf_folder = Path(pdf_folder)
    if not excel_path.is_file():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    if not pdf_folder.is_dir():
        raise FileNotFoundError(f"PDF folder not found: {pdf_folder}")

    copy_amount, copy_rule = lookup_excel_copy_amount(excel_path)
    emit(f"Excel supplementary count: {copy_amount} ({copy_rule})")

    rows = load_waybill_rows(excel_path)
    expected = {row.container for row in rows}
    emit(f"Loaded {len(rows)} container rows from Excel.")

    pdf_paths = sorted(
        (item for item in pdf_folder.iterdir() if item.is_file() and item.suffix.lower() == ".pdf"),
        key=lambda path: path.name.lower(),
    )
    if not pdf_paths:
        raise ValueError("No PDF files were found in the selected folder.")
    emit(f"Scanning {len(pdf_paths)} PDF files.")

    matches_by_container: dict[str, list[PdfMatch]] = {row.container: [] for row in rows}
    unmatched: list[PdfMatch] = []
    for index, pdf_path in enumerate(pdf_paths, start=1):
        match = identify_pdf(pdf_path, expected)
        if match.container:
            matches_by_container[match.container].append(match)
            emit(f"[{index}/{len(pdf_paths)}] {pdf_path.name} -> {match.container} ({match.method})")
        else:
            unmatched.append(match)
            emit(f"[{index}/{len(pdf_paths)}] {pdf_path.name} -> not matched")

    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    base_output = output_root if output_root is not None else pdf_folder
    output_dir = Path(base_output) / f"sorted_waybills_{timestamp}"
    sorted_dir = output_dir / "Sorted"
    unmatched_dir = output_dir / "Unmatched"
    sorted_dir.mkdir(parents=True, exist_ok=False)
    if unmatched:
        unmatched_dir.mkdir(parents=True, exist_ok=True)

    report_rows: list[list[str]] = []
    matched_count = 0
    missing_count = 0
    duplicate_count = 0
    verified_copy_count = 0
    verification_failure_count = 0
    sequence_totals: dict[int, int] = {}
    for row in rows:
        sequence_totals[row.sequence] = sequence_totals.get(row.sequence, 0) + 1

    for final_order, row in enumerate(rows, start=1):
        sequence_label = f"{row.sequence:03d}"
        if sequence_totals.get(row.sequence, 0) > 1:
            sequence_label = f"{sequence_label}-{row.sequence_occurrence:02d}"
        row_matches = matches_by_container[row.container]
        if not row_matches:
            missing_count += 1
            notes = "; ".join(item for item in (f"Excel row {row.excel_row}", row.sequence_note) if item)
            report_rows.append(
                [
                    str(final_order),
                    str(row.sequence),
                    str(row.sequence_occurrence),
                    row.container,
                    "missing",
                    "",
                    "",
                    notes,
                    "",
                    "",
                    "",
                ]
            )
            continue

        if len(row_matches) > 1:
            duplicate_count += len(row_matches) - 1

        for duplicate_index, match in enumerate(row_matches, start=1):
            duplicate_suffix = "" if duplicate_index == 1 else f"_DUP{duplicate_index}"
            output_name = f"{sequence_label}_{row.container}{duplicate_suffix}.pdf"
            output_path = unique_path(sorted_dir / output_name)
            shutil.copy2(match.path, output_path)
            source_hash = sha256_file(match.path)
            output_hash = sha256_file(output_path)
            verified_copy = source_hash == output_hash
            if verified_copy:
                verified_copy_count += 1
            else:
                verification_failure_count += 1
            matched_count += 1
            status = "matched" if duplicate_index == 1 else "duplicate"
            notes = "; ".join(item for item in (match.note, f"Excel row {row.excel_row}", row.sequence_note) if item)
            if not verified_copy:
                notes = "; ".join(item for item in (notes, "COPY VERIFICATION FAILED") if item)
            report_rows.append(
                [
                    str(final_order),
                    str(row.sequence),
                    str(row.sequence_occurrence),
                    row.container,
                    status,
                    str(match.path),
                    str(output_path),
                    notes,
                    source_hash,
                    output_hash,
                    "yes" if verified_copy else "no",
                ]
            )

    for match in unmatched:
        target = unique_path(unmatched_dir / match.path.name) if unmatched else match.path
        if unmatched:
            shutil.copy2(match.path, target)
        report_rows.append(["", "", "", "", match.method, str(match.path), str(target), match.note, "", "", ""])

    report_path = output_dir / "sort_report.csv"
    with report_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "sort_order",
                "sequence",
                "sequence_occurrence",
                "container",
                "status",
                "source_file",
                "output_file",
                "notes",
                "source_sha256",
                "output_sha256",
                "verified_copy",
            ]
        )
        writer.writerows(report_rows)

    summary_path = output_dir / "summary.txt"
    write_summary(
        summary_path,
        [
            "UniTrans Waybill Sort Summary",
            f"Excel: {excel_path}",
            f"PDF folder: {pdf_folder}",
            f"Copy amount: {copy_amount}",
            f"Copy rule: {copy_rule}",
            f"Expected containers: {len(rows)}",
            f"PDF files scanned: {len(pdf_paths)}",
            f"Matched PDFs copied: {matched_count}",
            f"Byte-for-byte verified copies: {verified_copy_count}/{matched_count}",
            f"Copy verification failures: {verification_failure_count}",
            f"Missing Excel containers: {missing_count}",
            f"Unmatched PDFs: {len(unmatched)}",
            f"Duplicate matched PDFs: {duplicate_count}",
            f"Sorted folder: {sorted_dir}",
            f"Report: {report_path}",
        ],
    )

    emit(f"Finished. Sorted files are in: {sorted_dir}")
    emit(f"Copy verification: {verified_copy_count}/{matched_count} sorted PDFs are byte-for-byte identical to the downloads.")
    if missing_count or unmatched:
        emit("Check sort_report.csv for missing or unmatched files.")
    if verification_failure_count:
        raise RuntimeError(
            f"Copy verification failed for {verification_failure_count} PDF(s). "
            f"Do not use the output. Check: {report_path}"
        )

    return SortResult(
        output_dir=output_dir,
        sorted_dir=sorted_dir,
        report_path=report_path,
        summary_path=summary_path,
        expected_count=len(rows),
        pdf_count=len(pdf_paths),
        matched_count=matched_count,
        missing_count=missing_count,
        unmatched_count=len(unmatched),
        duplicate_count=duplicate_count,
        verified_copy_count=verified_copy_count,
        verification_failure_count=verification_failure_count,
        copy_amount=copy_amount,
        copy_rule=copy_rule,
    )


def read_sorted_report(report_path: Path) -> list[tuple[int, str, Path]]:
    rows: list[tuple[int | None, int, str, Path]] = []
    with Path(report_path).open("r", newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        for record in reader:
            if record.get("status") not in {"matched", "duplicate"}:
                continue
            output_file = record.get("output_file") or ""
            sort_order = parse_sequence(record.get("sort_order"))
            sequence = parse_sequence(record.get("sequence"))
            container = normalize_container(record.get("container"))
            if sequence is None or not container or not output_file:
                continue
            rows.append((sort_order, sequence, container, Path(output_file)))

    rows.sort(key=lambda item: (item[0] is None, item[0] or item[1], item[3].name.lower()))
    return [(sequence, container, path) for _sort_order, sequence, container, path in rows]


def safe_copy_count(value: object, fallback: int = 1) -> int:
    parsed = parse_sequence(value)
    if parsed is None:
        return fallback
    return max(1, min(parsed, 99))


def infer_supplementary_count(page_count: int) -> int | None:
    # Current full downloads are 12 base pages plus 2 pages per supplementary copy.
    if page_count >= 12 and (page_count - 12) % 2 == 0:
        return (page_count - 12) // 2
    return None


def analyze_pdf_folder(pdf_folder: Path) -> PdfFolderAnalysis:
    import_dependencies_ok()

    counts: list[int] = []
    for pdf_path in sorted(Path(pdf_folder).glob("*.pdf"), key=lambda path: path.name.lower()):
        reader = PdfReader(str(pdf_path))
        if reader.is_encrypted:
            reader.decrypt("")
        counts.append(len(reader.pages))

    unique_counts = tuple(sorted(set(counts)))
    detected = infer_supplementary_count(unique_counts[0]) if len(unique_counts) == 1 and unique_counts else None
    return PdfFolderAnalysis(len(counts), unique_counts, detected)


def register_separator_font() -> str:
    font_name = "Helvetica"
    font_candidates = (
        Path("C:/Windows/Fonts/arial.ttf"),
        Path("C:/Windows/Fonts/calibri.ttf"),
        Path("C:/Windows/Fonts/msyh.ttc"),
        Path("C:/Windows/Fonts/simsun.ttc"),
    )
    for font_path in font_candidates:
        if not font_path.exists():
            continue
        try:
            name = f"SeparatorFont{abs(hash(str(font_path))) % 100000}"
            pdfmetrics.registerFont(TTFont(name, str(font_path)))
            return name
        except Exception:
            continue
    return font_name


def draw_logo(pdf: canvas.Canvas, page_width: float, page_height: float) -> None:
    path = logo_path()
    if not path or ImageReader is None:
        return
    try:
        image = ImageReader(str(path))
        image_width, image_height = image.getSize()
        target_width = page_width * 0.34
        target_height = target_width * image_height / image_width
        x = (page_width - target_width) / 2
        y = page_height * 0.70
        pdf.drawImage(image, x, y, width=target_width, height=target_height, preserveAspectRatio=True, mask="auto")
    except Exception:
        return


def build_separator_page(
    sequence: int,
    container: str,
    next_sequence: int | None = None,
    next_container: str | None = None,
) -> object:
    font_name = register_separator_font()
    buffer = io.BytesIO()
    page_width, page_height = A4
    pdf = canvas.Canvas(buffer, pagesize=A4)
    pdf.setTitle(f"{sequence:03d}_{container}_separator")
    draw_logo(pdf, page_width, page_height)
    pdf.setFillColorRGB(0.05, 0.16, 0.32)
    pdf.setFont(font_name, 40)
    pdf.drawCentredString(page_width / 2, page_height * 0.58, f"{sequence:03d} complete")
    pdf.setFont(font_name, 24)
    pdf.drawCentredString(page_width / 2, page_height * 0.50, container)
    pdf.setFont(font_name, 14)
    if next_sequence is not None and next_container:
        pdf.drawCentredString(page_width / 2, page_height * 0.43, f"Next: {next_sequence:03d}  {next_container}")
    else:
        pdf.drawCentredString(page_width / 2, page_height * 0.43, "End of batch")
    pdf.setStrokeColorRGB(0.85, 0.88, 0.92)
    pdf.line(page_width * 0.22, page_height * 0.36, page_width * 0.78, page_height * 0.36)
    pdf.setFont(font_name, 11)
    pdf.setFillColorRGB(0.35, 0.39, 0.45)
    pdf.drawCentredString(page_width / 2, page_height * 0.30, "Divider page")
    pdf.showPage()
    pdf.save()
    buffer.seek(0)
    return PdfReader(buffer).pages[0]


def should_stamp_sequence_page(page_text: str) -> bool:
    compact = re.sub(r"\s+", "", page_text or "")
    first_chunk = compact[:120]
    return (
        first_chunk.startswith("1") and "运单正本" in first_chunk
    ) or (
        first_chunk.startswith("4") and "运单副本" in first_chunk
    )


def build_sequence_stamp_overlay(page_width: float, page_height: float, sequence: int) -> object:
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=(page_width, page_height))
    label = f"{sequence:02d}" if sequence < 100 else str(sequence)
    x = page_width - 30
    y = 29
    pdf.setFont("Helvetica-Bold", 8)
    pdf.setFillColorRGB(0, 0, 0)
    pdf.drawString(x, y, label)
    pdf.save()
    buffer.seek(0)
    return PdfReader(buffer).pages[0]


def add_source_pdf_to_writer(
    writer: PdfWriter,
    reader: PdfReader,
    sequence: int,
    mark_sequence: bool,
) -> tuple[int, int]:
    page_count = len(reader.pages)
    stamped_count = 0
    if not mark_sequence:
        writer.append(reader)
        return page_count, stamped_count

    scan_limit = min(page_count, 12)
    for page_index, page in enumerate(reader.pages):
        should_check = stamped_count < 2 and page_index < scan_limit and page_index % 2 == 0
        if should_check and should_stamp_sequence_page(page.extract_text() or ""):
            width = float(page.mediabox.width)
            height = float(page.mediabox.height)
            page.merge_page(build_sequence_stamp_overlay(width, height, sequence))
            stamped_count += 1
        writer.add_page(page)
    return page_count, stamped_count


def create_print_batch_pdf(
    report_path: Path,
    output_dir: Path,
    include_separators: bool = False,
    mark_sequence: bool = False,
    log: Callable[[str], None] | None = None,
    progress: Callable[[int, int], None] | None = None,
) -> PrintBatchResult:
    import_dependencies_ok()

    def emit(message: str) -> None:
        if log:
            log(message)

    sorted_rows = read_sorted_report(report_path)
    if not sorted_rows:
        raise ValueError("No matched sorted PDFs were found in the report.")

    print_dir = Path(output_dir) / "Bulk PDF"
    print_dir.mkdir(parents=True, exist_ok=True)
    suffix = "_with_separators" if include_separators else ""
    output_path = unique_path(print_dir / f"bulk_waybills{suffix}.pdf")

    writer = PdfWriter()
    source_pages = 0
    output_pages = 0
    separator_pages = 0

    emit("Creating bulk PDF from sorted waybills.")
    if include_separators or mark_sequence:
        enabled = []
        if include_separators:
            enabled.append("separator pages")
        if mark_sequence:
            enabled.append("digital sequence marks")
        emit("Optional bulk PDF extras enabled: " + ", ".join(enabled) + ".")
    else:
        emit("Safe bulk PDF mode: original PDF pages only; no separator pages or digital sequence marks.")
    total_rows = len(sorted_rows)
    for row_index, (sequence, container, pdf_path) in enumerate(sorted_rows):
        if not pdf_path.is_file():
            raise FileNotFoundError(f"Sorted PDF not found: {pdf_path}")
        source_reader = PdfReader(str(pdf_path))
        if source_reader.is_encrypted:
            decrypt_result = source_reader.decrypt("")
            if decrypt_result == 0:
                raise ValueError(f"Could not open encrypted PDF: {pdf_path}")
        page_count, _stamped_count = add_source_pdf_to_writer(writer, source_reader, sequence, mark_sequence)
        source_pages += page_count
        output_pages += page_count

        if include_separators and row_index < len(sorted_rows) - 1:
            next_sequence, next_container, _ = sorted_rows[row_index + 1]
            writer.add_page(build_separator_page(sequence, container, next_sequence, next_container))
            separator_pages += 1
            output_pages += 1
        if progress:
            progress(row_index + 1, total_rows)

    with output_path.open("wb") as handle:
        writer.write(handle)

    emit(f"Bulk PDF created: {output_path}")
    return PrintBatchResult(
        path=output_path,
        waybill_count=len(sorted_rows),
        source_pages=source_pages,
        output_pages=output_pages,
        separator_pages=separator_pages,
    )

I18N = {
    "en": {
        "window_title": "Unitrans Portal",
        "brand": "UNITRANS",
        "workflow_title": "Choose Workflow",
        "workflow_subtitle": "Select what you want to do today.",
        "workflow_sort": "Sort Waybill PDFs",
        "workflow_sort_desc": "Sort downloaded waybill PDFs and create a safe bulk PDF.",
        "workflow_driver": "Fill Driver Numbers",
        "workflow_driver_desc": "Read scanned waybill images and fill blank car numbers in Excel.",
        "workflow_unlock": "Unlock Files",
        "workflow_unlock_desc": "Remove PDF restrictions and Excel workbook/sheet protection. Supports individual files or a folder.",
        "back_to_workflows": "Back to Main Menu",
        "hero_title": "Waybill Sorter",
        "hero_subtitle": "Sort waybills and create one bulk PDF in Excel order.",
        "status_start": "Choose an Excel file when you are ready.",
        "status_excel_loaded": "Excel loaded. Choose the PDF folder next.",
        "status_ready": "Ready to sort.",
        "status_sorting": "Sorting PDFs...",
        "status_creating_pdf": "Creating bulk PDF...",
        "status_finished": "Finished.",
        "status_driver_start": "Choose the Excel file for driver numbers.",
        "status_driver_excel_loaded": "Excel loaded. Choose the scan image folder next.",
        "status_driver_ready": "Ready to fill driver numbers.",
        "status_driver_working": "Filling driver numbers...",
        "status_unlock_start": "Choose files or a folder to unlock.",
        "status_unlock_ready": "Ready to unlock files.",
        "status_unlock_working": "Unlocking files...",
        "section_excel": "1. Excel",
        "section_pdf": "2. PDF Folder",
        "section_options": "3. Actions",
        "section_log": "Activity",
        "section_driver_excel": "1. Excel",
        "section_driver_images": "2. Scan Images",
        "section_unlock_pdf": "1. Files or Folder",
        "section_unlock_actions": "2. Actions",
        "driver_images": "Image files",
        "blank_car_numbers": "Blank car numbers",
        "choose_images": "Choose Image Folder",
        "choose_image_in_folder": "Select any image in the folder",
        "fill_driver_numbers": "Fill Driver Numbers",
        "driver_fill_finished": "Driver Numbers Filled",
        "driver_fill_failed": "Driver Fill Failed",
        "driver_fill_finished_msg": "Updated workbook:\n{workbook}\n\nFilled: {filled}\nAlready filled: {already}\nNeeds review: {review}\nUnmatched: {unmatched}",
        "missing_images_title": "Missing Image Folder",
        "missing_images": "Choose the folder containing the scanned waybill images.",
        "choose_excel": "Choose Excel",
        "excel_columns": "Excel Columns",
        "excel_columns_title": "Excel Columns",
        "excel_columns_note": "Use this when the workbook headers are unusual. Enter Excel letters such as A and C.",
        "header_row": "Header row",
        "sequence_column": "Sequence column",
        "container_column": "Container column",
        "apply": "Apply",
        "cancel": "Cancel",
        "choose_folder": "Choose PDF Folder",
        "choose_unlock_files": "Choose Files",
        "choose_unlock_folder": "Choose Folder",
        "unlock_password": "Password (if required)",
        "unlock_files_count": "Supported files",
        "choose_pdf_in_folder": "Select any PDF in the folder",
        "copy_amount": "Excel supplementary count",
        "container_rows": "Container rows",
        "pdf_files": "PDF files",
        "pdf_page_check": "Downloaded PDF check",
        "unlock_pdfs": "Unlock Files",
        "mark_sequence": "Digitally mark sequence on forms 1 and 4",
        "separators": "Add separator page between waybills",
        "sort": "Sort Waybills",
        "make_pdf": "Create Bulk PDF",
        "open_output": "Open Output Folder",
        "options": "Options",
        "options_title": "Optional PDF Extras",
        "options_note": "Leave these off for normal use. They change the bulk PDF only; sorted PDFs stay exact copies.",
        "close_options": "Close",
        "clear_log": "Clear Log",
        "language": "中文",
        "settings": "Settings",
        "settings_title": "Settings",
        "settings_about": "About",
        "settings_updates": "Updates",
        "settings_version": "Installed version",
        "settings_check_now": "Check for updates now",
        "settings_checking": "Checking for updates\u2026",
        "settings_up_to_date": "You are running the latest version.",
        "settings_update_available": "Version {version} is available.",
        "settings_check_failed": "Could not check for updates: {error}",
        "settings_auto_note": "This version does not check for updates on its own. Use the button above to check manually.",
        "settings_last_checked": "Last checked: {when}",
        "settings_never_checked": "Never checked",
        "settings_close": "Close",
        "settings_appearance": "Appearance",
        "settings_theme_light": "Light",
        "settings_theme_dark": "Dark",
        "theme_to_dark": "Dark mode",
        "theme_to_light": "Light mode",
        "update_title": "Update available",
        "update_heading": "Version {version} is available",
        "update_current": "You are running {current}.",
        "update_size": "Download size: {size}",
        "update_whats_new": "What's new",
        "update_download": "Open download page",
        "update_install": "Download and install",
        "update_install_restart": "Install and restart",
        "update_downloading": "Downloading update…",
        "update_progress": "Downloaded {received} of {total}",
        "update_ready": "Download verified. Install now to restart the app.",
        "update_installing": "Installing update and restarting…",
        "update_wait_for_work": "Finish the current task before installing the update.",
        "update_download_failed": "Update download failed: {error}",
        "update_download_cancelled": "Download cancelled. You can try again.",
        "update_install_failed": "Could not start the update: {error}",
        "update_later": "Later",
        "update_open_failed": "Could not open the download page in a browser.\n\n{url}",
        "update_notes_missing": "No release notes were published for this version.",
        "missing_excel_title": "Missing Excel",
        "missing_excel": "Choose the Excel file first.",
        "missing_folder_title": "Missing PDF Folder",
        "missing_folder": "Choose the folder containing the PDF waybills.",
        "missing_unlock_folder": "Choose one or more PDF/Excel files or a folder containing them.",
        "missing_sort_title": "Nothing Sorted Yet",
        "missing_sort": "Sort the waybills before creating a bulk PDF.",
        "excel_problem": "Excel Problem",
        "sort_failed": "Sort Failed",
        "pdf_failed": "Bulk PDF Failed",
        "unlock_failed": "File Unlock Failed",
        "sort_finished": "Sort Finished",
        "pdf_finished": "Bulk PDF Created",
        "unlock_finished": "File Unlock Finished",
        "sort_finished_msg": "Sorted folder:\n{sorted_dir}\n\nMatched PDFs: {matched}\nVerified exact copies: {verified}/{matched}\nMissing containers: {missing}\nUnmatched PDFs: {unmatched}",
        "pdf_finished_msg": "Bulk PDF:\n{path}\n\nPages: {pages}\nSeparator pages: {separators}\n\nOpen this PDF in Chrome or Adobe Reader to print.",
        "unlock_finished_msg": "Output folder:\n{output_dir}\n\nUnlocked: {unlocked}\nAlready unlocked copied: {already}\nSkipped password-required: {skipped}\nFailed: {failed}\n\nReport:\n{report}",
        "optional_pdf_features_warning_title": "Optional Extras Enabled",
        "optional_pdf_features_warning": "You enabled optional PDF extras. This changes the bulk PDF by adding separator pages and/or digital sequence marks.\n\nThe sorted PDFs remain byte-for-byte copies of the downloads.\n\nContinue creating the bulk PDF?",
        "yes": "Yes",
        "no": "No",
        "log_excel_loaded": "Excel loaded: {name}",
        "log_copy_amount": "Excel supplementary count: {amount} ({rule})",
        "log_container_rows": "Container rows: {count}",
        "log_pdf_folder": "PDF folder loaded: {path}",
        "log_pdf_count": "PDF files: {count}",
        "log_pdf_page_check": "Downloaded PDF check: {check}",
        "log_driver_excel": "Driver Excel loaded: {name}",
        "log_driver_rows": "Driver workflow rows: {rows}; blank car numbers: {blank}",
        "log_image_folder": "Image folder loaded: {path}",
        "log_image_count": "Image files: {count}",
        "log_starting_sort": "Starting sort.",
        "log_starting_driver_fill": "Starting driver number fill.",
        "log_starting_unlock": "Starting file unlock.",
        "log_output": "Output: {path}",
        "log_print_pdf": "Bulk PDF: {path}",
        "log_unlock_pdf_folder": "Unlock input loaded: {path}",
        "progress_print_pdf": "Creating bulk PDF: {percent}% ({current}/{total})",
        "progress_unlock_pdf": "Unlocking files: {percent}% ({current}/{total})",
        "log_copy_verification": "Copy verification: {verified}/{total} sorted PDFs are byte-for-byte identical to the downloads.",
        "pdf_check_unknown": "-",
        "pdf_check_uniform": "{pages} pages per PDF; detected supplementary count {detected}; Excel says {excel}; {status}",
        "pdf_check_mixed": "Mixed PDF page counts: {counts}",
        "pdf_check_ok": "OK",
        "pdf_check_mismatch": "check mismatch",
    },
    "zh": {
        "window_title": "Unitrans Portal 运单排序工具",
        "brand": "UNITRANS",
        "workflow_title": "选择工作流程",
        "workflow_subtitle": "请选择今天要处理的任务。",
        "workflow_sort": "排序运单 PDF",
        "workflow_sort_desc": "按照 Excel 顺序排序 PDF，并生成安全的合并 PDF。",
        "workflow_driver": "填写车号",
        "workflow_driver_desc": "读取运单扫描图片，并填写 Excel 中空白的车号。",
        "workflow_unlock": "\u89e3\u9501\u6587\u4ef6",
        "workflow_unlock_desc": "\u79fb\u9664 PDF \u9650\u5236\u4ee5\u53ca Excel \u5de5\u4f5c\u7c3f/\u5de5\u4f5c\u8868\u4fdd\u62a4\uff0c\u652f\u6301\u9009\u62e9\u6587\u4ef6\u6216\u6587\u4ef6\u5939\u3002",
        "back_to_workflows": "返回主菜单",
        "hero_title": "运单排序工具",
        "hero_subtitle": "按照 Excel 顺序整理运单，并生成一个合并 PDF。",
        "status_start": "准备好后请选择 Excel 文件。",
        "status_excel_loaded": "Excel 已读取。下一步请选择 PDF 文件夹。",
        "status_ready": "已准备好，可以开始排序。",
        "status_sorting": "正在排序 PDF...",
        "status_creating_pdf": "正在生成合并 PDF...",
        "status_finished": "完成。",
        "status_driver_start": "请选择需要填写车号的 Excel 文件。",
        "status_driver_excel_loaded": "Excel 已读取。下一步请选择扫描图片文件夹。",
        "status_driver_ready": "已准备好，可以填写车号。",
        "status_driver_working": "正在填写车号...",
        "status_unlock_start": "\u8bf7\u9009\u62e9\u8981\u89e3\u9501\u7684\u6587\u4ef6\u6216\u6587\u4ef6\u5939\u3002",
        "status_unlock_ready": "\u5df2\u51c6\u5907\u597d\uff0c\u53ef\u4ee5\u89e3\u9501\u6587\u4ef6\u3002",
        "status_unlock_working": "\u6b63\u5728\u89e3\u9501\u6587\u4ef6...",
        "section_excel": "1. Excel",
        "section_pdf": "2. PDF 文件夹",
        "section_options": "3. 操作",
        "section_log": "操作记录",
        "section_driver_excel": "1. Excel",
        "section_driver_images": "2. 扫描图片",
        "section_unlock_pdf": "1. \u6587\u4ef6\u6216\u6587\u4ef6\u5939",
        "section_unlock_actions": "2. \u64cd\u4f5c",
        "driver_images": "图片文件数",
        "blank_car_numbers": "空白车号",
        "choose_images": "选择图片文件夹",
        "choose_image_in_folder": "请选择该文件夹中的任意图片",
        "fill_driver_numbers": "填写车号",
        "driver_fill_finished": "车号填写完成",
        "driver_fill_failed": "车号填写失败",
        "driver_fill_finished_msg": "更新后的工作簿：\n{workbook}\n\n已填写：{filled}\n已有车号：{already}\n建议复查：{review}\n未匹配：{unmatched}",
        "missing_images_title": "缺少图片文件夹",
        "missing_images": "请选择包含运单扫描图片的文件夹。",
        "choose_excel": "选择 Excel",
        "excel_columns": "Excel 列设置",
        "excel_columns_title": "Excel 列设置",
        "excel_columns_note": "当工作表的表头不常见时使用。请填写 Excel 列字母，例如 A 和 C。",
        "header_row": "表头行",
        "sequence_column": "序号列",
        "container_column": "箱号列",
        "apply": "应用",
        "cancel": "取消",
        "choose_folder": "选择 PDF 文件夹",
        "choose_unlock_files": "选择文件",
        "choose_unlock_folder": "选择文件夹",
        "unlock_password": "密码（如需要）",
        "unlock_files_count": "支持的文件",
        "choose_pdf_in_folder": "请选择该文件夹中的任意 PDF",
        "copy_amount": "Excel 补充文件数",
        "container_rows": "箱号行数",
        "pdf_files": "PDF 文件数",
        "pdf_page_check": "下载 PDF 检查",
        "unlock_pdfs": "\u89e3\u9501\u6587\u4ef6",
        "mark_sequence": "在第 1 和第 4 联添加电子序号",
        "separators": "每票之间加入分隔页",
        "sort": "排序运单",
        "make_pdf": "生成合并 PDF",
        "open_output": "打开输出文件夹",
        "options": "选项",
        "options_title": "可选 PDF 功能",
        "options_note": "正常使用请保持关闭。这些只会改变合并 PDF；排序后的 PDF 仍是原文件的精确副本。",
        "close_options": "关闭",
        "clear_log": "清空记录",
        "language": "English",
        "settings": "设置",
        "settings_title": "设置",
        "settings_about": "关于",
        "settings_updates": "更新",
        "settings_version": "已安装版本",
        "settings_check_now": "立即检查更新",
        "settings_checking": "正在检查更新…",
        "settings_up_to_date": "当前已是最新版本。",
        "settings_update_available": "有新版本 {version} 可用。",
        "settings_check_failed": "无法检查更新：{error}",
        "settings_auto_note": "此版本不会自动检查更新。请使用上方按钮手动检查。",
        "settings_last_checked": "上次检查：{when}",
        "settings_never_checked": "从未检查",
        "settings_close": "关闭",
        "settings_appearance": "外观",
        "settings_theme_light": "浅色",
        "settings_theme_dark": "深色",
        "theme_to_dark": "深色模式",
        "theme_to_light": "浅色模式",
        "update_title": "有可用更新",
        "update_heading": "有新版本 {version} 可用",
        "update_current": "当前版本为 {current}。",
        "update_size": "下载大小：{size}",
        "update_whats_new": "更新内容",
        "update_download": "打开下载页面",
        "update_install": "下载并安装",
        "update_install_restart": "安装并重启",
        "update_downloading": "正在下载更新…",
        "update_progress": "已下载 {received} / {total}",
        "update_ready": "下载已验证。现在安装将重新启动应用。",
        "update_installing": "正在安装更新并重新启动…",
        "update_wait_for_work": "请先完成当前任务，再安装更新。",
        "update_download_failed": "更新下载失败：{error}",
        "update_download_cancelled": "已取消下载。您可以稍后重试。",
        "update_install_failed": "无法启动更新：{error}",
        "update_later": "稍后",
        "update_open_failed": "无法在浏览器中打开下载页面。\n\n{url}",
        "update_notes_missing": "此版本没有发布更新说明。",
        # See the English block: the footer version label is language-neutral.
        "missing_excel_title": "缺少 Excel",
        "missing_excel": "请先选择 Excel 文件。",
        "missing_folder_title": "缺少 PDF 文件夹",
        "missing_folder": "请选择包含 PDF 运单的文件夹。",
        "missing_unlock_folder": "\u8bf7\u9009\u62e9\u4e00\u4e2a\u6216\u591a\u4e2a PDF/Excel \u6587\u4ef6\uff0c\u6216\u5305\u542b\u76f8\u5173\u6587\u4ef6\u7684\u6587\u4ef6\u5939\u3002",
        "missing_sort_title": "尚未排序",
        "missing_sort": "请先排序运单，再生成合并 PDF。",
        "excel_problem": "Excel 问题",
        "sort_failed": "排序失败",
        "pdf_failed": "合并 PDF 生成失败",
        "unlock_failed": "\u6587\u4ef6\u89e3\u9501\u5931\u8d25",
        "sort_finished": "排序完成",
        "pdf_finished": "合并 PDF 已生成",
        "unlock_finished": "\u6587\u4ef6\u89e3\u9501\u5b8c\u6210",
        "sort_finished_msg": "排序文件夹：\n{sorted_dir}\n\n已匹配 PDF：{matched}\n精确副本校验：{verified}/{matched}\n缺少箱号：{missing}\n未匹配 PDF：{unmatched}",
        "pdf_finished_msg": "合并 PDF：\n{path}\n\n页数：{pages}\n分隔页：{separators}\n\n请用 Chrome 或 Adobe Reader 打开后打印。",
        "unlock_finished_msg": "\u8f93\u51fa\u6587\u4ef6\u5939\uff1a\n{output_dir}\n\n\u5df2\u89e3\u9501\uff1a{unlocked}\n\u5df2\u53ef\u7528\u5e76\u5df2\u590d\u5236\uff1a{already}\n\u9700\u8981\u6253\u5f00\u5bc6\u7801\u5df2\u8df3\u8fc7\uff1a{skipped}\n\u5931\u8d25\uff1a{failed}\n\n\u62a5\u544a\uff1a\n{report}",
        "optional_pdf_features_warning_title": "已开启可选功能",
        "optional_pdf_features_warning": "你开启了可选 PDF 功能。合并 PDF 会加入分隔页和/或电子序号。\n\n排序后的 PDF 仍是下载文件的逐字节副本。\n\n继续生成合并 PDF 吗？",
        "yes": "是",
        "no": "否",
        "log_excel_loaded": "已读取 Excel：{name}",
        "log_copy_amount": "Excel 补充文件数：{amount} ({rule})",
        "log_container_rows": "箱号行数：{count}",
        "log_pdf_folder": "已选择 PDF 文件夹：{path}",
        "log_pdf_count": "PDF 文件数：{count}",
        "log_pdf_page_check": "下载 PDF 检查：{check}",
        "log_driver_excel": "车号 Excel 已读取：{name}",
        "log_driver_rows": "车号流程行数：{rows}；空白车号：{blank}",
        "log_image_folder": "已选择图片文件夹：{path}",
        "log_image_count": "图片文件数：{count}",
        "log_starting_sort": "开始排序。",
        "log_starting_driver_fill": "开始填写车号。",
        "log_starting_unlock": "\u5f00\u59cb\u89e3\u9501\u6587\u4ef6\u3002",
        "log_output": "输出：{path}",
        "log_print_pdf": "合并 PDF：{path}",
        "log_unlock_pdf_folder": "\u5df2\u9009\u62e9\u89e3\u9501\u8f93\u5165\uff1a{path}",
        "progress_print_pdf": "正在生成合并 PDF：{percent}%（{current}/{total}）",
        "progress_unlock_pdf": "\u6b63\u5728\u89e3\u9501\u6587\u4ef6\uff1a{percent}%\uff08{current}/{total}\uff09",
        "log_copy_verification": "复制校验：{verified}/{total} 个排序 PDF 与下载文件逐字节一致。",
        "pdf_check_unknown": "-",
        "pdf_check_uniform": "每个 PDF {pages} 页；检测补充文件数 {detected}；Excel 显示 {excel}；{status}",
        "pdf_check_mixed": "PDF 页数不一致：{counts}",
        "pdf_check_ok": "正常",
        "pdf_check_mismatch": "请检查不一致",
    },
}


class WaybillSorterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.language = tk.StringVar(value="en")
        self.tk.call("tk", "scaling", 1.25)
        self.title(self.tr("window_title"))
        self.geometry("1180x760")
        self.minsize(1040, 660)
        self.configure(bg=BG_WINDOW)

        self.excel_path = tk.StringVar()
        self.excel_column_override: tuple[int, int, int] | None = None
        self.pdf_folder = tk.StringVar()
        self.copy_amount = tk.StringVar(value="-")
        self.container_count = tk.StringVar(value="-")
        self.pdf_count = tk.StringVar(value="-")
        self.pdf_page_check = tk.StringVar(value="-")
        self.include_separators = tk.BooleanVar(value=False)
        self.mark_sequence = tk.BooleanVar(value=False)
        self.progress_value = tk.DoubleVar(value=0)
        self.progress_text = tk.StringVar(value="")
        self.status_key = "status_start"
        self.status_text = tk.StringVar(value=self.tr("status_start"))
        self.worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker_running = False
        self.latest_result: SortResult | None = None
        self.latest_print_batch: PrintBatchResult | None = None
        self.latest_print_signature: tuple[str, bool] | None = None
        self.current_workflow = "selector"
        self.driver_excel_path = tk.StringVar()
        self.driver_image_folder = tk.StringVar()
        self.driver_row_count = tk.StringVar(value="-")
        self.driver_blank_count = tk.StringVar(value="-")
        self.driver_image_count = tk.StringVar(value="-")
        self.latest_driver_result: DriverFillResult | None = None
        self.unlock_file_paths: tuple[Path, ...] = ()
        self.unlock_source_label = tk.StringVar()
        self.unlock_password = tk.StringVar()
        self.unlock_pdf_count = tk.StringVar(value="-")
        self.latest_unlock_result: PdfUnlockResult | None = None
        self.options_window: tk.Toplevel | None = None
        self.settings_window: tk.Toplevel | None = None
        # Loaded once at startup; written back when the user changes something.
        self.user_settings: dict[str, object] = app_settings.load_settings()

        # Establish the colour scheme before any widget is created, so the first
        # window is drawn in the saved scheme rather than flashing light and
        # then switching.
        apply_theme(str(self.user_settings.get("theme") or LIGHT))
        self.configure(bg=BG_WINDOW)

        # Owns the update-check thread. Kept so a close during a check does not
        # leave a callback pointing at a destroyed window.
        self.update_check_token = 0
        self.update_check_running = False
        self.log_history: list[tuple[str, str]] = []
        self.widgets: dict[str, object] = {}

        self._build_workflow_selector()

    def tr(self, key: str, **kwargs: object) -> str:
        text = I18N[self.language.get()].get(key, I18N["en"].get(key, key))
        return text.format(**kwargs) if kwargs else text

    def _card(self, parent: tk.Widget, title: str) -> tuple[tk.Frame, tk.Frame]:
        shell = tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        title_label = tk.Label(
            shell,
            text=title,
            bg=SURFACE,
            fg=TEXT_HEADING,
            font=("Segoe UI", 13, "bold"),
            anchor="w",
        )
        title_label.pack(fill=tk.X, padx=16, pady=(12, 2))
        body = tk.Frame(shell, bg=SURFACE)
        body.pack(fill=tk.BOTH, expand=True, padx=16, pady=(8, 14))
        return shell, body

    def _primary_button(
        self,
        parent: tk.Widget,
        text: str,
        command: Callable[[], None],
        compact: bool = False,
        label_key: str | None = None,
    ) -> tk.Button:
        uniform = self._uniform_width(label_key)
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=ACCENT,
            activebackground=ACCENT_HOVER,
            fg=ON_ACCENT,
            activeforeground=ON_ACCENT,
            disabledforeground=ON_ACCENT,
            relief=tk.FLAT,
            bd=0,
            # highlightthickness is a Tk default of 1-2px and draws the focus
            # ring. Left unset it shows as a stray blue box around a button
            # whenever that button takes keyboard focus - which happens on the
            # first click in a rebuilt window, making it look like the language
            # switch changed something. takefocus=0 keeps it out of the tab
            # order so the ring never appears.
            highlightthickness=0,
            takefocus=0,
            padx=18 if compact else 22,
            pady=9 if compact else 12,
            # A shared width keeps sibling buttons equal and stops the layout
            # resizing when the language changes.
            width=uniform if uniform else 0,
            font=("Segoe UI", 11 if compact else 12, "bold"),
            cursor="hand2",
        )

    def _set_primary_button_enabled(self, button: tk.Button, enabled: bool) -> None:
        button.configure(
            state=tk.NORMAL if enabled else tk.DISABLED,
            bg=ACCENT if enabled else "#6F9FCE",
            activebackground=ACCENT_HOVER if enabled else "#6F9FCE",
            fg=ON_ACCENT,
            activeforeground=ON_ACCENT,
            disabledforeground=ON_ACCENT,
            cursor="hand2" if enabled else "arrow",
        )

    def _button_exists(self, name: str) -> bool:
        widget = getattr(self, name, None)
        try:
            return bool(widget is not None and widget.winfo_exists())
        except tk.TclError:
            return False

    def _action_button_width(self, compact: bool = False) -> int:
        """One pixel width that fits every button label, in every language.

        Without this each button sizes to its own text, so they come out
        different widths - and because the longest label differs per language,
        whole columns resize when the language is switched. Measuring the
        widest label across all languages and reusing that width keeps the
        buttons equal and stops the layout shifting.

        The full-size and compact (path-row) buttons are measured against both
        fonts so neither category can clip, which also makes every button in
        the app the same width as every other. Cached, because the measurement
        does not depend on the active language.
        """
        cached = getattr(self, "_action_width_cache", None)
        if cached is not None:
            return cached

        try:
            from tkinter import font as tkfont

            fonts = [
                tkfont.Font(root=self, font=("Segoe UI", 12, "bold")),
                tkfont.Font(root=self, font=("Segoe UI", 11, "bold")),
            ]
            widest = 0
            for language in I18N.values():
                for key in ACTION_BUTTON_KEYS:
                    text = str(language.get(key, ""))
                    if not text:
                        continue
                    for font in fonts:
                        widest = max(
                            widest,
                            font.measure(text),
                            # The back button carries a leading arrow.
                            font.measure(f"\u2190  {text}"),
                        )
            # 22px padding each side in the factory, plus slack so a font
            # substitution on another machine cannot clip the longest label.
            width = widest + 44 + 8
        except Exception:
            width = 0

        width = max(160, min(width, 340))
        self._action_width_cache = width
        return width

    def _uniform_width(self, key: str | None) -> int | None:
        """Width for a button labelled ``key``, in Tk's button width units.

        Tk's ``width`` option on a button counts *characters*, not pixels, and
        Tk treats it as a **minimum**, not a fixed size: a button whose label is
        wider still grows, and one whose label is narrower is *not* padded out
        to it. So this alone cannot make a row of buttons match - the shorter
        labels come out narrower. It sets the floor; ``_stack_button`` supplies
        the fill that makes every button in a column the same width.

        Primary and secondary buttons use different font sizes, so each is
        calibrated separately; combining the floor here with the fill there is
        what makes them come out the same number of pixels wide.
        """
        if key not in ACTION_BUTTON_KEYS:
            return None

        pixel_target = self._action_button_width()
        try:
            from tkinter import font as tkfont

            # Calibrate against 100 "M"s so the per-character average is stable.
            probe = tkfont.Font(root=self, font=("Segoe UI", 12, "bold"))
            per_char = max(1.0, probe.measure("M" * 100) / 100.0)
        except Exception:
            per_char = 11.0

        units = max(8, int(round(pixel_target / per_char)))
        return units

    def _stack_button(
        self,
        parent: tk.Widget,
        row: int,
        text: str,
        command: Callable[[], None],
        label_key: str | None = None,
        primary: bool = True,
        compact: bool = False,
        pady: tuple[int, int] = (0, 8),
    ) -> tk.Button:
        """Place a button in a vertical stack, stretched to the stack's width.

        Each button takes its own grid row with ``sticky="ew"`` while the
        stack's single column carries the weight, so every button adopts the
        width of the widest one in that stack. Together with the minimum from
        ``_uniform_width`` this is what makes a column of buttons genuinely
        equal in size.
        """
        parent.columnconfigure(0, weight=1)
        button = (
            self._primary_button(parent, text, command, compact=compact, label_key=label_key)
            if primary
            else self._secondary_button(parent, text, command, label_key=label_key)
        )
        button.grid(row=row, column=0, sticky="ew", pady=pady)
        return button

    def _dialog_button(
        self,
        parent: tk.Widget,
        text: str,
        command: Callable[[], None],
        label_key: str | None = None,
        primary: bool = True,
    ) -> tk.Button:
        """A full-width button in a dialog.

        Dialogs lay their contents out with ``pack``, and Tk refuses to mix
        ``pack`` and ``grid`` in one container: using the grid-based
        ``_stack_button`` here raises "cannot use geometry manager grid inside
        ... which already has slaves managed by pack". This is the pack
        equivalent.
        """
        button = (
            self._primary_button(parent, text, command, label_key=label_key)
            if primary
            else self._secondary_button(parent, text, command, label_key=label_key)
        )
        button.pack(fill=tk.X)
        return button

    def _secondary_button(
        self,
        parent: tk.Widget,
        text: str,
        command: Callable[[], None],
        label_key: str | None = None,
    ) -> tk.Button:
        uniform = self._uniform_width(label_key)
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=ACCENT_SOFT,
            activebackground=ACCENT_SOFT_HOVER,
            fg=TEXT_HEADING,
            activeforeground=TEXT_HEADING,
            relief=tk.FLAT,
            bd=0,
            # See _primary_button: without these a focus ring appears as a blue
            # box around the button.
            highlightthickness=0,
            takefocus=0,
            # 17px keeps the total height and text inset in step with a primary
            # button so the shared width produces equal-sized buttons.
            padx=17,
            pady=9,
            width=uniform if uniform else 0,
            font=("Segoe UI", 11, "bold"),
            cursor="hand2",
        )

    def _header_control(
        self,
        parent: tk.Widget,
        key: str,
        text: str,
        command: Callable[[], None],
        theme_button: bool = False,
    ) -> tk.Frame:
        """Put a changing header label in a fixed-size pixel box."""
        from tkinter import font as tkfont

        variants = [str(language.get(key, "")) for language in I18N.values()]
        if theme_button:
            variants = [
                f"{icon}  {language.get(label_key, '')}"
                for language in I18N.values()
                for icon, label_key in (("☀", "theme_to_light"), ("☾", "theme_to_dark"))
            ]
        label_font = tkfont.Font(root=self, font=("Segoe UI", 11, "bold"))
        width = max((label_font.measure(label) for label in variants), default=0) + 48
        shell = tk.Frame(parent, bg=SURFACE, width=width, height=44)
        # The button is packed into this frame, so pack propagation is the
        # setting that matters here. grid_propagate(False) leaves the packer
        # free to resize the shell to the translated button's requested size.
        shell.pack_propagate(False)
        button = (
            self._theme_button(shell)
            if theme_button
            else self._secondary_button(shell, text, command)
        )
        if theme_button:
            self.theme_button = button
        button.configure(width=0)
        button.pack(fill=tk.BOTH, expand=True)
        return shell

    def _back_button(self, parent: tk.Widget) -> tk.Button:
        """A clearly visible 'back to the main menu' button.

        Deliberately larger and bolder than a secondary button: it was
        previously styled the same as 'Options' and 'Clear Log' and was easy to
        miss, which is why it read as missing.
        """
        return tk.Button(
            parent,
            text=f"\u2190  {self.tr('back_to_workflows')}",
            command=self.return_to_workflows,
            bg=ACCENT,
            activebackground=ACCENT_HOVER,
            fg=ON_ACCENT,
            activeforeground=ON_ACCENT,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            takefocus=0,
            padx=20,
            pady=11,
            width=self._uniform_width("back_to_workflows") or 0,
            font=("Segoe UI", 11, "bold"),
            cursor="hand2",
        )

    def load_logo_image(self, max_width: int = 270, max_height: int = 86) -> object | None:
        if Image is None or ImageTk is None:
            return None
        path = logo_path()
        if not path:
            return None
        try:
            image = Image.open(path)
            image.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
            self.logo_image = ImageTk.PhotoImage(image)
            return self.logo_image
        except Exception:
            return None

    def _path_row(
        self,
        parent: tk.Widget,
        button_text: str,
        command: Callable[[], None],
        variable: tk.StringVar,
        label_key: str | None = None,
    ) -> None:
        button = self._primary_button(
            parent, button_text, command, compact=True, label_key=label_key
        )
        # Width 0 removes the character-count floor. Tk's width option is a
        # *minimum*, and for this smaller font it lands above the shared target
        # - which would make these buttons wider than the rest of the column.
        # The column is given the target as a minimum size instead and the
        # button stretched across it, so they match the taller action buttons
        # exactly whatever the font.
        button.configure(width=0)
        target = self._action_button_width() if self._uniform_width(label_key) else 0
        button.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        if target:
            parent.columnconfigure(0, minsize=target, weight=0)
        entry = tk.Entry(
            parent,
            textvariable=variable,
            state="readonly",
            readonlybackground=BG_LIST,
            fg=TEXT_ENTRY,
            relief=tk.FLAT,
            font=("Segoe UI", 11),
        )
        entry.grid(row=0, column=1, sticky="ew", ipady=12)
        parent.columnconfigure(1, weight=1)

    def _stat(
        self,
        parent: tk.Widget,
        row: int,
        label: str,
        variable: tk.StringVar,
        label_key: str | None = None,
    ) -> None:
        variants = [
            str(language[label_key])
            for language in I18N.values()
            if label_key is not None and label_key in language
        ] or [label]
        try:
            from tkinter import font as tkfont

            label_font = tkfont.Font(root=self, font=("Segoe UI", 10))
            widest = max(label_font.measure(text) for text in variants)
            char_width = max(1, label_font.measure("0"))
            stable_width = (widest + char_width - 1) // char_width
        except Exception:
            stable_width = 0
        tk.Label(
            parent,
            text=label,
            bg=SURFACE,
            fg=TEXT_LABEL,
            font=("Segoe UI", 10),
            width=stable_width,
            anchor="w",
        ).grid(
            row=row,
            column=0,
            sticky="w",
            pady=(8, 0),
        )
        tk.Label(parent, textvariable=variable, bg=SURFACE, fg=TEXT, font=("Segoe UI", 11, "bold"), wraplength=440, justify="left").grid(
            row=row,
            column=1,
            sticky="w",
            pady=(8, 0),
            padx=(10, 0),
        )

    def _clear_window(self) -> None:
        for child in self.winfo_children():
            child.destroy()

    def _configure_base_style(self) -> None:
        """Theme the ttk widgets, which ignore Tk's colour options.

        ``ttk.Scrollbar`` and ``ttk.Separator`` are drawn by the ttk theme
        engine, not by Tk's own colour options, so assigning ``bg``/``fg`` to
        them does nothing. They have to be configured through a style - and
        left alone they stay light grey in dark mode, which is exactly the kind
        of detail that makes a dark theme look unfinished.
        """
        style = ttk.Style(self)
        try:
            # "clam" is the only built-in theme that honours these settings
            # across platforms; the Windows native theme ignores them.
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TSpinbox", arrowsize=14)
        style.configure("TSeparator", background=DIVIDER)
        style.configure(
            "TScrollbar",
            background=SCROLLBAR,
            troughcolor=SCROLLBAR_TROUGH,
            bordercolor=SCROLLBAR_TROUGH,
            arrowcolor=TEXT_MUTED,
            lightcolor=SCROLLBAR,
            darkcolor=SCROLLBAR,
        )
        style.map(
            "TScrollbar",
            background=[("active", ACCENT_SOFT_HOVER), ("pressed", ACCENT_HOVER)],
        )

    def _build_header(self, root: tk.Widget, title: str, subtitle: str, back: bool = False) -> None:
        """Header strip: optionally a back button, then the title, then actions.

        The logo is deliberately *not* shown inside a workflow. It is a wide
        bitmap that crowds the title, and it pushed the back button and the
        heading into adjacent columns with barely any gap between them. Dropping
        it inside a workflow gives the back button and the title room to
        breathe, and puts the back button at the very left edge where it is
        easiest to find. The logo still heads the workflow selector.
        """
        hero = tk.Frame(root, bg=SURFACE)
        hero.pack(fill=tk.X)
        hero.columnconfigure(1, weight=1)

        title_column = 0
        if back:
            self._back_button(hero).grid(
                row=0, column=0, sticky="w", padx=(18, 0), pady=14
            )
            title_column = 1
        else:
            logo_image = self.load_logo_image()
            # The logo is navy artwork with a transparent background, and its
            # strapline is dark blue. Dropped straight onto a dark surface the
            # strapline all but disappears, so in dark mode the mark keeps a
            # light plate behind it and stays legible as supplied.
            logo_bg = LIGHT_LOGO_PLATE if theme.active() == DARK else SURFACE
            brand = tk.Label(hero, bg=logo_bg, padx=12, pady=6)
            if logo_image is not None:
                brand.configure(image=logo_image)
            else:
                brand.configure(
                    text=self.tr("brand"),
                    bg=logo_bg,
                    fg=TEXT_STRONG,
                    font=("Segoe UI", 24, "bold"),
                )
            brand.grid(row=0, column=0, sticky="w", padx=(18, 0), pady=10)
            title_column = 1

        title_frame = tk.Frame(hero, bg=SURFACE)
        title_frame.grid(
            row=0,
            column=title_column,
            sticky="ew",
            # Generous separation from the back button: they are unrelated
            # controls, so they should not read as one block.
            padx=(32 if back else 24, 12),
            pady=12,
        )
        tk.Label(
            title_frame,
            text=title,
            bg=SURFACE,
            fg=TEXT_STRONG,
            font=("Segoe UI", 21, "bold"),
            anchor="w",
        ).pack(fill=tk.X)
        tk.Label(
            title_frame,
            text=subtitle,
            bg=SURFACE,
            fg=TEXT_MUTED,
            font=("Segoe UI", 11),
            anchor="w",
        ).pack(fill=tk.X, pady=(2, 0))
        self._header_control(hero, "settings", self.tr("settings"), self.open_settings).grid(
            row=0,
            column=4,
            sticky="e",
            padx=(0, 18),
        )
        self._header_control(hero, "theme", "", self.toggle_theme, theme_button=True).grid(
            row=0, column=3, sticky="e", padx=(0, 14)
        )
        self._header_control(hero, "language", self.tr("language"), self.toggle_language).grid(
            row=0,
            column=2,
            sticky="e",
            padx=(0, 14),
        )

    def _theme_button(self, parent: tk.Widget) -> tk.Button:
        """Toggle between the light and dark schemes.

        Labelled with the scheme the button switches *to*, which is what a user
        expects from a toggle, and shows a sun or a moon so the current state is
        legible at a glance without reading the label.
        """
        dark_now = theme.active() == DARK
        label = self.tr("theme_to_light") if dark_now else self.tr("theme_to_dark")
        icon = "\u2600" if dark_now else "\u263d"  # sun when dark, moon when light
        return tk.Button(
            parent,
            text=f"{icon}  {label}",
            command=self.toggle_theme,
            bg=ACCENT_SOFT,
            activebackground=ACCENT_SOFT_HOVER,
            fg=TEXT_HEADING,
            activeforeground=TEXT_HEADING,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            takefocus=0,
            padx=17,
            pady=9,
            font=("Segoe UI", 11, "bold"),
            cursor="hand2",
        )

    def _build_footer(self, root: tk.Widget) -> None:
        footer = tk.Frame(root, bg=BG_WINDOW, padx=16, pady=0)
        footer.pack(fill=tk.X)
        tk.Label(
            footer,
            text=version_label(),
            bg=BG_WINDOW,
            fg=TEXT_FAINT,
            font=("Segoe UI", 9),
            anchor="e",
        ).pack(fill=tk.X, pady=4)

    def _build_log_panel(self, parent: tk.Widget) -> None:
        log_shell, log_body = self._card(parent, self.tr("section_log"))
        log_shell.grid(row=0, column=1, sticky="nsew")
        log_body.rowconfigure(0, weight=1)
        log_body.columnconfigure(0, weight=1)
        self.log_text = tk.Text(
            log_body,
            height=14,
            wrap=tk.WORD,
            font=("Consolas", 12),
            bg=BG_PANEL,
            fg=TEXT,
            relief=tk.FLAT,
            padx=12,
            pady=12,
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_body, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.render_log_history()

    def _workflow_choice(
        self,
        parent: tk.Widget,
        row: int,
        title: str,
        description: str,
        command: Callable[[], None],
        button_key: str | None = None,
    ) -> None:
        shell = tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        shell.grid(row=row, column=0, sticky="ew", pady=(0, 14))
        # The text column absorbs spare width; the button column keeps the
        # fixed button width, so the three cards line up and the layout does
        # not shift when the language changes.
        shell.columnconfigure(0, weight=1)
        tk.Label(
            shell,
            text=title,
            bg=SURFACE,
            fg=TEXT_HEADING,
            font=("Segoe UI", 16, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 4))
        tk.Label(
            shell,
            text=description,
            bg=SURFACE,
            fg=TEXT_MUTED,
            font=("Segoe UI", 11),
            anchor="w",
            justify="left",
            wraplength=700,
        ).grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 16))
        self._primary_button(shell, title, command, label_key=button_key).grid(
            row=0, column=1, rowspan=2, sticky="nsew", padx=18, pady=16
        )

    def _build_workflow_selector(self) -> None:
        self.current_workflow = "selector"
        self._configure_base_style()
        root = tk.Frame(self, bg=BG_WINDOW)
        root.pack(fill=tk.BOTH, expand=True)
        self._build_header(root, self.tr("workflow_title"), self.tr("workflow_subtitle"))

        body = tk.Frame(root, bg=BG_WINDOW, padx=48, pady=38)
        body.pack(fill=tk.BOTH, expand=True)
        body.columnconfigure(0, weight=1)
        choices = tk.Frame(body, bg=BG_WINDOW)
        choices.grid(row=0, column=0, sticky="nsew")
        choices.columnconfigure(0, weight=1)
        self._workflow_choice(
            choices,
            0,
            self.tr("workflow_sort"),
            self.tr("workflow_sort_desc"),
            self.select_sort_workflow,
            button_key="workflow_sort",
        )
        self._workflow_choice(
            choices,
            1,
            self.tr("workflow_driver"),
            self.tr("workflow_driver_desc"),
            self.select_driver_workflow,
            button_key="workflow_driver",
        )
        self._workflow_choice(
            choices,
            2,
            self.tr("workflow_unlock"),
            self.tr("workflow_unlock_desc"),
            self.select_unlock_workflow,
            button_key="workflow_unlock",
        )
        self._build_footer(root)

    def select_sort_workflow(self) -> None:
        if self.worker_running:
            return
        self.status_key = "status_start"
        self.status_text.set(self.tr(self.status_key))
        self._clear_window()
        self._build_ui()
        if self.pdf_folder.get():
            self.update_pdf_folder_analysis(Path(self.pdf_folder.get()), log=False)

    def select_driver_workflow(self) -> None:
        if self.worker_running:
            return
        self.status_key = "status_driver_start"
        self.status_text.set(self.tr(self.status_key))
        self._clear_window()
        self._build_driver_ui()

    def select_unlock_workflow(self) -> None:
        if self.worker_running:
            return
        self.status_key = "status_unlock_start"
        self.status_text.set(self.tr(self.status_key))
        self._clear_window()
        self._build_unlock_ui()

    def _build_unlock_ui(self) -> None:
        self.current_workflow = "unlock"
        self._configure_base_style()

        root = tk.Frame(self, bg=BG_WINDOW)
        root.pack(fill=tk.BOTH, expand=True)
        self._build_header(root, self.tr("workflow_unlock"), self.tr("workflow_unlock_desc"), back=True)

        outer = tk.Frame(root, bg=BG_WINDOW, padx=16, pady=14)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        folder_shell, folder_body = self._card(outer, self.tr("section_unlock_pdf"))
        folder_shell.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        chooser_row = tk.Frame(folder_body, bg=SURFACE)
        chooser_row.grid(row=0, column=0, columnspan=2, sticky="ew")
        chooser_row.columnconfigure(2, weight=1)
        for column, key, command in (
            (0, "choose_unlock_files", self.choose_unlock_files),
            (1, "choose_unlock_folder", self.choose_unlock_folder),
        ):
            button = self._primary_button(
                chooser_row, self.tr(key), command, compact=True, label_key=key
            )
            button.configure(width=0)
            button.grid(row=0, column=column, sticky="ew", padx=(0, 8))
            chooser_row.columnconfigure(column, minsize=self._action_button_width())
        source_entry = tk.Entry(
            chooser_row, textvariable=self.unlock_source_label, state="readonly",
            readonlybackground=BG_LIST, fg=TEXT_ENTRY, relief=tk.FLAT,
            font=("Segoe UI", 11),
        )
        source_entry.grid(row=0, column=2, sticky="ew", ipady=12)
        password_row = tk.Frame(folder_body, bg=SURFACE)
        password_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        tk.Label(
            password_row, text=self.tr("unlock_password"), bg=SURFACE,
            fg=TEXT_LABEL, font=("Segoe UI", 10),
        ).pack(side=tk.LEFT, padx=(0, 12))
        tk.Entry(
            password_row, textvariable=self.unlock_password, show="*",
            bg=BG_LIST, fg=TEXT_ENTRY, relief=tk.FLAT, font=("Segoe UI", 11),
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
        self._stat(folder_body, 2, self.tr("unlock_files_count"), self.unlock_pdf_count, "unlock_files_count")

        lower = tk.Frame(outer, bg=BG_WINDOW)
        lower.grid(row=1, column=0, sticky="nsew")
        lower.columnconfigure(0, weight=0)
        lower.columnconfigure(1, weight=1)
        lower.rowconfigure(0, weight=1)

        actions_shell, actions_body = self._card(lower, self.tr("section_unlock_actions"))
        actions_shell.grid(row=0, column=0, sticky="nsw", padx=(0, 14))
        action_stack = tk.Frame(actions_body, bg=SURFACE)
        action_stack.grid(row=0, column=0, sticky="ew")
        self.unlock_button = self._stack_button(
            action_stack,
            1,
            self.tr("unlock_pdfs"),
            self.start_unlock_pdfs,
            label_key="unlock_pdfs",
        )
        self.unlock_open_output_button = self._stack_button(
            action_stack,
            2,
            self.tr("open_output"),
            self.open_unlock_output,
            label_key="open_output",
        )
        self._stack_button(
            action_stack,
            3,
            self.tr("clear_log"),
            self.clear_log,
            label_key="clear_log",
            primary=False,
            pady=(0, 0),
        )

        self._build_log_panel(lower)
        self.update_action_state()
        self._build_footer(root)

    def _build_driver_ui(self) -> None:
        self.current_workflow = "driver"
        self._configure_base_style()

        root = tk.Frame(self, bg=BG_WINDOW)
        root.pack(fill=tk.BOTH, expand=True)
        self._build_header(root, self.tr("workflow_driver"), self.tr("workflow_driver_desc"), back=True)

        outer = tk.Frame(root, bg=BG_WINDOW, padx=16, pady=14)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        top_grid = tk.Frame(outer, bg=BG_WINDOW)
        top_grid.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        top_grid.columnconfigure(0, weight=1)
        top_grid.columnconfigure(1, weight=1)

        excel_shell, excel_body = self._card(top_grid, self.tr("section_driver_excel"))
        excel_shell.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._path_row(
            excel_body, self.tr("choose_excel"), self.choose_driver_excel,
            self.driver_excel_path, label_key="choose_excel",
        )
        self._stat(excel_body, 1, self.tr("container_rows"), self.driver_row_count, "container_rows")
        self._stat(excel_body, 2, self.tr("blank_car_numbers"), self.driver_blank_count, "blank_car_numbers")

        image_shell, image_body = self._card(top_grid, self.tr("section_driver_images"))
        image_shell.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        self._path_row(
            image_body, self.tr("choose_images"), self.choose_driver_image_folder,
            self.driver_image_folder, label_key="choose_images",
        )
        self._stat(image_body, 1, self.tr("driver_images"), self.driver_image_count, "driver_images")

        lower = tk.Frame(outer, bg=BG_WINDOW)
        lower.grid(row=1, column=0, sticky="nsew")
        lower.columnconfigure(0, weight=0)
        lower.columnconfigure(1, weight=1)
        lower.rowconfigure(0, weight=1)

        actions_shell, actions_body = self._card(lower, self.tr("section_options"))
        actions_shell.grid(row=0, column=0, sticky="nsw", padx=(0, 14))
        action_stack = tk.Frame(actions_body, bg=SURFACE)
        action_stack.grid(row=0, column=0, sticky="ew")
        self.driver_fill_button = self._stack_button(
            action_stack,
            4,
            self.tr("fill_driver_numbers"),
            self.start_driver_fill,
            label_key="fill_driver_numbers",
        )
        self.driver_open_output_button = self._stack_button(
            action_stack,
            5,
            self.tr("open_output"),
            self.open_driver_output,
            label_key="open_output",
        )
        self._stack_button(
            action_stack,
            6,
            self.tr("clear_log"),
            self.clear_log,
            label_key="clear_log",
            primary=False,
            pady=(0, 0),
        )

        self._build_log_panel(lower)
        self.update_action_state()
        self._build_footer(root)

    def return_to_workflows(self) -> None:
        if self.worker_running:
            return
        self.status_key = "status_start"
        self.status_text.set(self.tr(self.status_key))
        self._clear_window()
        self._build_workflow_selector()

    def _build_ui(self) -> None:
        self.current_workflow = "sort"
        self._configure_base_style()

        root = tk.Frame(self, bg=BG_WINDOW)
        root.pack(fill=tk.BOTH, expand=True)

        self._build_header(root, self.tr("hero_title"), self.tr("hero_subtitle"), back=True)

        outer = tk.Frame(root, bg=BG_WINDOW, padx=16, pady=14)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        top_grid = tk.Frame(outer, bg=BG_WINDOW)
        top_grid.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        top_grid.columnconfigure(0, weight=1)
        top_grid.columnconfigure(1, weight=1)

        excel_shell, excel_body = self._card(top_grid, self.tr("section_excel"))
        excel_shell.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._path_row(
            excel_body, self.tr("choose_excel"), self.choose_excel,
            self.excel_path, label_key="choose_excel",
        )
        self._stat(excel_body, 1, self.tr("copy_amount"), self.copy_amount, "copy_amount")
        self._stat(excel_body, 2, self.tr("container_rows"), self.container_count, "container_rows")
        # Sits below the two stat rows, in column 0 - the same column as the
        # "Choose Excel" picker above it, so the two should end at the same x
        # position. Equal width in Tk's units is not enough to guarantee that:
        # Tk grows a button past its requested width when the label needs the
        # room, and these two labels differ in length. The button is therefore
        # wrapped in a frame that is matched to the picker's rendered width once
        # the first layout pass has run - the only reliable way to make two
        # differently-labelled buttons agree.
        excel_columns_button = self._secondary_button(
            excel_body, self.tr("excel_columns"), self.open_excel_columns_dialog,
            label_key="excel_columns",
        )
        excel_columns_button.configure(width=0)
        holder = tk.Frame(excel_body, bg=SURFACE)
        holder.grid(row=3, column=0, sticky="w", pady=(14, 0))
        excel_columns_button.pack(in_=holder, fill=tk.X)

        picker = next(
            (
                child
                for child in excel_body.winfo_children()
                if isinstance(child, tk.Button)
            ),
            None,
        )

        def match_picker_width() -> None:
            """Give the holder the picker's exact width, once it is known."""
            try:
                if picker is None or not holder.winfo_exists():
                    return
                width = picker.winfo_width()
                # A width of 1 means the layout has not run yet.
                if width > 1:
                    holder.configure(width=width)
                    holder.pack_propagate(False)
            except tk.TclError:
                pass

        try:
            self.after(80, match_picker_width)
        except tk.TclError:
            pass

        pdf_shell, pdf_body = self._card(top_grid, self.tr("section_pdf"))
        pdf_shell.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        self._path_row(
            pdf_body, self.tr("choose_folder"), self.choose_pdf_folder,
            self.pdf_folder, label_key="choose_folder",
        )
        self._stat(pdf_body, 1, self.tr("pdf_files"), self.pdf_count, "pdf_files")
        self._stat(pdf_body, 2, self.tr("pdf_page_check"), self.pdf_page_check, "pdf_page_check")

        lower = tk.Frame(outer, bg=BG_WINDOW)
        lower.grid(row=1, column=0, sticky="nsew")
        lower.columnconfigure(0, weight=0)
        lower.columnconfigure(1, weight=1)
        lower.rowconfigure(0, weight=1)

        options_shell, options_body = self._card(lower, self.tr("section_options"))
        options_shell.grid(row=0, column=0, sticky="nsw", padx=(0, 14))
        action_stack = tk.Frame(options_body, bg=SURFACE)
        action_stack.grid(row=0, column=0, sticky="ew")
        self.sort_button = self._stack_button(
            action_stack,
            7,
            self.tr("sort"),
            self.start_sort,
            label_key="sort",
        )
        self.make_pdf_button = self._stack_button(
            action_stack,
            8,
            self.tr("make_pdf"),
            self.create_print_pdf,
            label_key="make_pdf",
        )
        self.open_output_button = self._stack_button(
            action_stack,
            9,
            self.tr("open_output"),
            self.open_output,
            label_key="open_output",
        )
        self.options_button = self._stack_button(
            action_stack,
            10,
            self.tr("options"),
            self.open_options,
            label_key="options",
            primary=False,
        )
        self._stack_button(
            action_stack,
            11,
            self.tr("clear_log"),
            self.clear_log,
            label_key="clear_log",
            primary=False,
            pady=(0, 0),
        )

        log_shell, log_body = self._card(lower, self.tr("section_log"))
        log_shell.grid(row=0, column=1, sticky="nsew")
        log_body.rowconfigure(0, weight=1)
        log_body.columnconfigure(0, weight=1)
        self.log_text = tk.Text(
            log_body,
            height=14,
            wrap=tk.WORD,
            font=("Consolas", 12),
            bg=BG_PANEL,
            fg=TEXT,
            relief=tk.FLAT,
            padx=12,
            pady=12,
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_body, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.update_action_state()
        self.render_log_history()

        footer = tk.Frame(root, bg=BG_WINDOW, padx=16, pady=0)
        footer.pack(fill=tk.X)
        tk.Label(
            footer,
            text=version_label(),
            bg=BG_WINDOW,
            fg=TEXT_FAINT,
            font=("Segoe UI", 9),
            anchor="e",
        ).pack(fill=tk.X, pady=4)

    def set_progress(self, value: float = 0, text: str = "") -> None:
        self.progress_value.set(max(0, min(100, value)))
        self.progress_text.set(text)

    def format_pdf_page_check(self, analysis: PdfFolderAnalysis) -> str:
        if not analysis.page_counts:
            return self.tr("pdf_check_unknown")
        if len(analysis.page_counts) > 1:
            return self.tr("pdf_check_mixed", counts=", ".join(str(count) for count in analysis.page_counts))

        pages = analysis.page_counts[0]
        detected = analysis.detected_supplementary_count
        excel_amount = parse_sequence(self.copy_amount.get())
        if detected is None:
            return f"{pages} pages per PDF; supplementary count unknown"
        if excel_amount is None:
            status = "-"
            excel_display = "-"
        else:
            status = self.tr("pdf_check_ok") if excel_amount == detected else self.tr("pdf_check_mismatch")
            excel_display = str(excel_amount)
        return self.tr(
            "pdf_check_uniform",
            pages=pages,
            detected=detected,
            excel=excel_display,
            status=status,
        )

    def update_pdf_folder_analysis(self, folder: Path, log: bool = True) -> None:
        try:
            analysis = analyze_pdf_folder(folder)
        except Exception:
            self.pdf_page_check.set(self.tr("pdf_check_unknown"))
            return
        self.pdf_count.set(str(analysis.pdf_count))
        check = self.format_pdf_page_check(analysis)
        self.pdf_page_check.set(check)
        if log:
            self.append_log(f"PDF files: {analysis.pdf_count}")
            self.append_log(f"Downloaded PDF check: {check}")

    def render_log_message(self, message: str) -> str:
        if self.language.get() != "zh":
            return message
        windows_patterns = (
            (r"^Windows OCR completed: (.+)/(.+) image\(s\) returned text\.$", "Windows OCR 已完成：{0}/{1} 张图片返回文本。"),
            (r"^Windows OCR skipped (.+) image\(s\); built-in image OCR will handle them\.$", "Windows OCR 跳过 {0} 张图片；将使用内置图片 OCR 处理。"),
            (r"^Windows OCR unavailable; using built-in image OCR only\. (.+)$", "Windows OCR 不可用；仅使用内置图片 OCR。{0}"),
            (r"^Windows OCR result could not be read; using built-in image OCR only\. (.+)$", "Windows OCR 结果无法读取；仅使用内置图片 OCR。{0}"),
            (r"^Windows OCR container matches: (.+)\.$", "Windows OCR 箱号匹配：{0}。"),
        )
        for pattern, replacement in windows_patterns:
            match = re.match(pattern, message)
            if match:
                return replacement.format(*match.groups())
        return self.localize_log_v2(message)

    def append_log(self, message: str, timestamp: str | None = None, record: bool = True) -> None:
        timestamp = timestamp or dt.datetime.now().strftime("%H:%M:%S")
        if record:
            self.log_history.append((timestamp, message))
        # The log widget only exists on screens that show a log panel; the
        # workflow selector does not. Anything logged before then is still kept
        # in log_history and rendered by render_log_history when the panel
        # appears.
        log_text = getattr(self, "log_text", None)
        if log_text is None:
            return
        try:
            if not log_text.winfo_exists():
                return
            log_text.insert(tk.END, f"[{timestamp}] {self.render_log_message(message)}\n")
            log_text.see(tk.END)
        except tk.TclError:
            pass

    def render_log_history(self) -> None:
        if not hasattr(self, "log_text"):
            return
        self.log_text.delete("1.0", tk.END)
        for timestamp, message in self.log_history:
            self.append_log(message, timestamp=timestamp, record=False)

    def open_options(self) -> None:
        if self.worker_running:
            return
        if self.options_window is not None:
            try:
                if self.options_window.winfo_exists():
                    self.options_window.lift()
                    self.options_window.focus_force()
                    return
            except tk.TclError:
                self.options_window = None

        window = tk.Toplevel(self)
        self.options_window = window
        window.title(self.tr("options_title"))
        window.configure(bg=SURFACE)
        window.resizable(False, False)
        window.transient(self)
        # Keep the dialog's outer frame constant across both languages; its
        # labels and button captions change, but Windows never resizes it.
        window.geometry("540x680")

        body = tk.Frame(window, bg=SURFACE, padx=22, pady=20)
        body.pack(fill=tk.BOTH, expand=True)
        tk.Label(
            body,
            text=self.tr("options_note"),
            bg=SURFACE,
            fg=TEXT_MUTED,
            font=("Segoe UI", 11),
            justify="left",
            wraplength=440,
        ).pack(fill=tk.X, pady=(0, 16))
        tk.Checkbutton(
            body,
            text=self.tr("separators"),
            variable=self.include_separators,
            bg=SURFACE,
            activebackground=SURFACE,
            fg=TEXT_HEADING,
            font=("Segoe UI", 11),
            anchor="w",
            selectcolor=SURFACE,
        ).pack(fill=tk.X, pady=(0, 10))
        tk.Checkbutton(
            body,
            text=self.tr("mark_sequence"),
            variable=self.mark_sequence,
            bg=SURFACE,
            activebackground=SURFACE,
            fg=TEXT_HEADING,
            font=("Segoe UI", 11),
            anchor="w",
            selectcolor=SURFACE,
        ).pack(fill=tk.X, pady=(0, 18))
        close_button = self._dialog_button(
            body,
            self.tr("close_options"),
            window.destroy,
            label_key="close_options",
        )

        def on_close() -> None:
            self.options_window = None
            window.destroy()

        close_button.configure(command=on_close)
        window.protocol("WM_DELETE_WINDOW", on_close)
        window.update_idletasks()
        x = self.winfo_rootx() + max(40, (self.winfo_width() - window.winfo_width()) // 2)
        y = self.winfo_rooty() + max(40, (self.winfo_height() - window.winfo_height()) // 3)
        window.geometry(f"+{x}+{y}")

    def open_settings(self) -> None:
        """Settings window: version information and a manual update check."""
        if self.settings_window is not None:
            try:
                if self.settings_window.winfo_exists():
                    self.settings_window.lift()
                    self.settings_window.focus_force()
                    return
            except tk.TclError:
                self.settings_window = None

        window = tk.Toplevel(self)
        self.settings_window = window
        window.title(self.tr("settings_title"))
        window.configure(bg=SURFACE)
        window.resizable(False, False)
        window.transient(self)

        body = tk.Frame(window, bg=SURFACE, padx=24, pady=20)
        body.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            body, text=self.tr("settings_about"), bg=SURFACE, fg=TEXT_HEADING,
            font=("Segoe UI", 12, "bold"), anchor="w",
        ).pack(fill=tk.X)
        tk.Label(
            body, text=self.tr("settings_version"), bg=SURFACE, fg=TEXT_MUTED,
            font=("Segoe UI", 10), anchor="w",
        ).pack(fill=tk.X, pady=(8, 0))
        tk.Label(
            body, text=version_label(), bg=SURFACE, fg=TEXT_STRONG,
            font=("Segoe UI", 13, "bold"), anchor="w",
        ).pack(fill=tk.X)

        ttk.Separator(body, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=16)

        # --- appearance ----------------------------------------------------
        tk.Label(
            body, text=self.tr("settings_appearance"), bg=SURFACE, fg=TEXT_HEADING,
            font=("Segoe UI", 12, "bold"), anchor="w",
        ).pack(fill=tk.X)

        scheme = theme.active()
        scheme_var = tk.StringVar(value=scheme)
        for value, label_key in ((LIGHT, "settings_theme_light"), (DARK, "settings_theme_dark")):
            tk.Radiobutton(
                body,
                text=self.tr(label_key),
                variable=scheme_var,
                value=value,
                command=lambda chosen=value: self.set_theme(chosen),
                bg=SURFACE,
                activebackground=SURFACE,
                fg=TEXT_HEADING,
                activeforeground=TEXT_HEADING,
                selectcolor=SURFACE,
                font=("Segoe UI", 10),
                anchor="w",
                highlightthickness=0,
                takefocus=0,
            ).pack(fill=tk.X, pady=(6 if value == LIGHT else 2, 0))

        ttk.Separator(body, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=16)

        tk.Label(
            body, text=self.tr("settings_updates"), bg=SURFACE, fg=TEXT_HEADING,
            font=("Segoe UI", 12, "bold"), anchor="w",
        ).pack(fill=tk.X)

        status = tk.Label(
            body, text=self._last_checked_text(), bg=SURFACE, fg=TEXT_MUTED,
            font=("Segoe UI", 10), anchor="w", justify="left", wraplength=430,
        )
        status.pack(fill=tk.X, pady=(8, 10))

        check_button = self._primary_button(
            body, self.tr("settings_check_now"), lambda: None, label_key="settings_check_now"
        )
        check_button.configure(command=lambda: start_check())
        check_button.pack(fill=tk.X)

        tk.Label(
            body, text=self.tr("settings_auto_note"), bg=SURFACE, fg=TEXT_FAINT,
            font=("Segoe UI", 9), anchor="w", justify="left", wraplength=430,
        ).pack(fill=tk.X, pady=(10, 16))

        close_button = self._dialog_button(
            body,
            self.tr("settings_close"),
            window.destroy,
            label_key="settings_close",
            primary=False,
        )

        def start_check() -> None:
            if self.update_check_running:
                return
            self.update_check_running = True
            self._set_primary_button_enabled(check_button, False)
            status.configure(text=self.tr("settings_checking"), fg=TEXT_MUTED)

            # Bumped whenever a new check starts; a result carrying a stale
            # token is discarded rather than written into a recycled window.
            self.update_check_token += 1
            token = self.update_check_token

            def worker() -> None:
                result = check_for_update()
                # Hop back to the Tk main thread - widgets must not be touched
                # from a worker.
                try:
                    self.after(0, lambda: finish(result, token))
                except (tk.TclError, RuntimeError):
                    pass  # window went away while the request was in flight

            def finish(result, check_token: int) -> None:
                if check_token != self.update_check_token:
                    return
                self.update_check_running = False
                try:
                    if not window.winfo_exists():
                        return
                except tk.TclError:
                    return
                self._set_primary_button_enabled(check_button, True)
                self._apply_update_check_result(result, status)

            threading.Thread(target=worker, name="update-check", daemon=True).start()

        def on_close() -> None:
            self.settings_window = None
            window.destroy()

        window.protocol("WM_DELETE_WINDOW", on_close)
        window.update_idletasks()
        x = self.winfo_rootx() + max(40, (self.winfo_width() - window.winfo_width()) // 2)
        y = self.winfo_rooty() + max(40, (self.winfo_height() - window.winfo_height()) // 3)
        window.geometry(f"+{x}+{y}")

    def set_theme(self, name: str) -> None:
        """Apply a specific scheme, as chosen in Settings."""
        if name == theme.active():
            return
        old_tokens = _load_colour_tokens()
        self.user_settings = app_settings.update_settings(theme=name)
        apply_theme(name)
        self.configure(bg=BG_WINDOW)
        self._refresh_current_screen(old_tokens=old_tokens)

    def _last_checked_text(self) -> str:
        """'Last checked: ...' or 'Never checked'."""
        stored = str(self.user_settings.get("last_update_check") or "")
        if not stored:
            return self.tr("settings_never_checked")
        try:
            moment = dt.datetime.fromisoformat(stored)
        except ValueError:
            return self.tr("settings_never_checked")
        return self.tr("settings_last_checked", when=moment.strftime("%Y-%m-%d %H:%M"))

    def _apply_update_check_result(self, result: object, status: tk.Label) -> None:
        """Render the outcome of an update check into the settings window."""
        if not getattr(result, "ok", False):
            status.configure(
                text=self.tr("settings_check_failed", error=getattr(result, "error", "")),
                fg=STATE_ERROR,
            )
            return

        # Record the successful check so "last checked" is meaningful next time.
        self.user_settings = app_settings.update_settings(
            last_update_check=dt.datetime.now().isoformat(timespec="seconds")
        )

        if getattr(result, "update_available", False):
            remote = str(getattr(result, "remote_version", ""))
            status.configure(
                text=self.tr("settings_update_available", version=remote),
                fg=ACCENT_TEXT,
            )
            self.show_update_dialog(result)
        else:
            status.configure(
                text=f"{self.tr('settings_up_to_date')}  ({self._last_checked_text()})",
                fg=STATE_SUCCESS,
            )

    def show_update_dialog(self, result: object) -> None:
        """Show release notes and safely download/apply a packaged update."""
        remote = str(getattr(result, "remote_version", ""))
        window = tk.Toplevel(self)
        window.title(self.tr("update_title"))
        window.configure(bg=SURFACE)
        window.transient(self)

        body = tk.Frame(window, bg=SURFACE, padx=24, pady=20)
        body.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            body, text=self.tr("update_heading", version=remote), bg=SURFACE,
            fg=TEXT_STRONG, font=("Segoe UI", 15, "bold"), anchor="w",
        ).pack(fill=tk.X)
        tk.Label(
            body, text=self.tr("update_current", current=version_label()), bg=SURFACE,
            fg=TEXT_MUTED, font=("Segoe UI", 10), anchor="w",
        ).pack(fill=tk.X, pady=(4, 0))

        size_text = format_size(getattr(result, "size_bytes", None))
        if size_text:
            tk.Label(
                body, text=self.tr("update_size", size=size_text), bg=SURFACE,
                fg=TEXT_MUTED, font=("Segoe UI", 10), anchor="w",
            ).pack(fill=tk.X, pady=(2, 0))

        tk.Label(
            body, text=self.tr("update_whats_new"), bg=SURFACE, fg=TEXT_HEADING,
            font=("Segoe UI", 12, "bold"), anchor="w",
        ).pack(fill=tk.X, pady=(16, 6))

        language = "zh" if self.language.get() == "zh" else "en"
        notes = result.notes_for(language) if hasattr(result, "notes_for") else []
        if not notes:
            notes = [self.tr("update_notes_missing")]

        notes_frame = tk.Frame(body, bg=SURFACE)
        notes_frame.pack(fill=tk.BOTH, expand=True)
        notes_text = tk.Text(
            notes_frame, height=9, wrap=tk.WORD, font=("Segoe UI", 10),
            bg=BG_PANEL, fg=TEXT, relief=tk.FLAT, padx=10, pady=10,
            highlightbackground=BORDER, highlightthickness=1,
        )
        notes_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(notes_frame, orient=tk.VERTICAL, command=notes_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        notes_text.configure(yscrollcommand=scrollbar.set)
        for note in notes:
            notes_text.insert(tk.END, f"\u2022 {note}\n\n")
        notes_text.configure(state=tk.DISABLED)

        progress_label = tk.Label(
            body, text="", bg=SURFACE, fg=TEXT_MUTED, font=("Segoe UI", 9),
            anchor="w", justify="left", wraplength=500,
        )
        progress_bar = ttk.Progressbar(body, orient=tk.HORIZONTAL, mode="determinate")

        buttons = tk.Frame(body, bg=SURFACE)
        buttons.pack(fill=tk.X, pady=(16, 0))

        url = str(getattr(result, "release_page", "") or getattr(result, "download_url", ""))
        cancel_event = threading.Event()
        state: dict[str, object] = {"busy": False, "staged_path": None, "install_started": False}

        def open_download() -> None:
            try:
                if not webbrowser.open(url):
                    raise RuntimeError("browser refused")
            except Exception:
                messagebox.showerror(
                    self.tr("update_title"),
                    self.tr("update_open_failed", url=url),
                    parent=window,
                )

        def close_dialog() -> None:
            cancel_event.set()
            if state.get("install_started"):
                self.destroy()
                return
            staged = state.get("staged_path")
            if staged and not state.get("install_started"):
                try:
                    Path(staged).unlink(missing_ok=True)
                except OSError:
                    pass
            try:
                window.destroy()
            except tk.TclError:
                pass

        def update_progress(received: int, total: int | None) -> None:
            try:
                self.after(0, lambda: render_progress(received, total))
            except (tk.TclError, RuntimeError):
                pass

        def render_progress(received: int, total: int | None) -> None:
            try:
                if not window.winfo_exists():
                    return
                if total:
                    progress_bar.configure(mode="determinate", maximum=total, value=received)
                    progress_label.configure(
                        text=self.tr(
                            "update_progress",
                            received=format_size(received),
                            total=format_size(total),
                        )
                    )
                else:
                    progress_label.configure(text=self.tr("update_downloading"))
            except tk.TclError:
                pass

        def download_failed(error: Exception) -> None:
            state["busy"] = False
            try:
                if not window.winfo_exists():
                    return
                progress_bar.stop()
                progress_label.configure(
                    text=(
                        self.tr("update_download_cancelled")
                        if isinstance(error, UpdateCancelled)
                        else self.tr("update_download_failed", error=str(error))
                    ),
                    fg=STATE_ERROR if not isinstance(error, UpdateCancelled) else TEXT_MUTED,
                )
                self._set_primary_button_enabled(install_button, True)
                install_button.configure(command=start_download)
            except tk.TclError:
                pass

        def download_complete(staged_path: Path) -> None:
            state["busy"] = False
            state["staged_path"] = staged_path
            try:
                if not window.winfo_exists():
                    staged_path.unlink(missing_ok=True)
                    return
                progress_bar.stop()
                expected_size = getattr(result, "size_bytes", None) or 100
                progress_bar.configure(mode="determinate", maximum=expected_size, value=expected_size)
                progress_label.configure(text=self.tr("update_ready"), fg=STATE_SUCCESS)
                install_button.configure(text=self.tr("update_install_restart"), command=install_update)
                self._set_primary_button_enabled(install_button, True)
            except tk.TclError:
                staged_path.unlink(missing_ok=True)

        def start_download() -> None:
            if state["busy"]:
                return
            if not supports_self_update():
                open_download()
                return
            if self.worker_running:
                progress_label.pack(fill=tk.X, pady=(10, 4), before=buttons)
                progress_label.configure(text=self.tr("update_wait_for_work"), fg=STATE_WARNING)
                return

            state["busy"] = True
            state["staged_path"] = None
            cancel_event.clear()
            progress_label.pack(fill=tk.X, pady=(10, 4), before=buttons)
            progress_label.configure(text=self.tr("update_downloading"), fg=TEXT_MUTED)
            progress_bar.pack(fill=tk.X, pady=(0, 2), before=buttons)
            expected_size = getattr(result, "size_bytes", None)
            if expected_size:
                progress_bar.stop()
                progress_bar.configure(mode="determinate", maximum=expected_size, value=0)
            else:
                progress_bar.configure(mode="indeterminate")
                progress_bar.start(12)
            self._set_primary_button_enabled(install_button, False)

            def worker() -> None:
                try:
                    staged_path = download_update(
                        result,
                        progress_callback=update_progress,
                        cancel_event=cancel_event,
                    )
                    self.after(0, lambda: download_complete(staged_path))
                except Exception as exc:
                    try:
                        self.after(0, lambda error=exc: download_failed(error))
                    except (tk.TclError, RuntimeError):
                        pass

            threading.Thread(target=worker, name="update-download", daemon=True).start()

        def install_update() -> None:
            staged_path = state.get("staged_path")
            if not staged_path:
                return
            if self.worker_running:
                progress_label.configure(text=self.tr("update_wait_for_work"), fg=STATE_WARNING)
                return
            try:
                launch_update_helper(
                    Path(staged_path),
                    str(getattr(result, "sha256", "")),
                )
            except UpdateDownloadError as exc:
                progress_label.configure(
                    text=self.tr("update_install_failed", error=str(exc)), fg=STATE_ERROR
                )
                install_button.configure(text=self.tr("update_install_restart"), command=install_update)
                self._set_primary_button_enabled(install_button, True)
                return

            state["install_started"] = True
            progress_label.configure(text=self.tr("update_installing"), fg=TEXT_MUTED)
            self._set_primary_button_enabled(install_button, False)
            later_button.configure(state=tk.DISABLED)
            # The helper is already waiting on this process; exit cleanly so it
            # can replace the executable, then relaunch the verified build.
            self.after(250, self.destroy)

        initial_install_label = self.tr("update_install") if supports_self_update() else self.tr("update_download")
        install_button = self._primary_button(
            buttons,
            initial_install_label,
            start_download,
            label_key="update_install" if supports_self_update() else "update_download",
        )
        if supports_self_update():
            try:
                from tkinter import font as tkfont

                button_font = tkfont.Font(root=self, font=install_button.cget("font"))
                labels = [
                    str(language[key])
                    for language in I18N.values()
                    for key in ("update_install", "update_install_restart")
                ]
                char_width = max(1, button_font.measure("0"))
                width = (max(button_font.measure(label) for label in labels) + char_width - 1) // char_width + 2
                install_button.configure(width=width)
            except Exception:
                pass
        install_button.pack(side=tk.RIGHT)
        later_button = self._secondary_button(buttons, self.tr("update_later"), close_dialog)
        later_button.pack(side=tk.RIGHT, padx=(0, 8))
        window.protocol("WM_DELETE_WINDOW", close_dialog)

        window.update_idletasks()
        x = self.winfo_rootx() + max(40, (self.winfo_width() - window.winfo_width()) // 2)
        y = self.winfo_rooty() + max(40, (self.winfo_height() - window.winfo_height()) // 3)
        window.geometry(f"+{x}+{y}")

    def localize_log_v2(self, message: str) -> str:
        patterns = (
            (r"^Copy amount from Excel name: (.+)$", "Excel 名称识别补充文件数：{0}"),
            (r"^Excel loaded: (.+)$", "已读取 Excel：{0}"),
            (r"^Excel supplementary count: (.+)$", "Excel 补充文件数：{0}"),
            (r"^Container rows: (.+)$", "箱号行数：{0}"),
            (r"^PDF folder loaded: (.+)$", "已选择 PDF 文件夹：{0}"),
            (r"^Unlock input loaded: (.+)$", "已选择解锁输入：{0}"),
            (r"^PDF files: (.+)$", "PDF 文件数：{0}"),
            (r"^Downloaded PDF check: (.+)$", "下载 PDF 检查：{0}"),
            (r"^Starting sort\.$", "开始排序。"),
            (r"^Starting file unlock\.$", "开始解锁文件。"),
            (r"^Driver Excel loaded: (.+)$", "车号 Excel 已读取：{0}"),
            (r"^Driver workflow rows: (.+); blank car numbers: (.+)$", "车号流程行数：{0}；空白车号：{1}"),
            (r"^Rows already containing car numbers: (.+)$", "已有车号的行数：{0}"),
            (r"^Image folder loaded: (.+)$", "已选择图片文件夹：{0}"),
            (r"^Image files: (.+)$", "图片文件数：{0}"),
            (r"^Starting driver number fill\.$", "开始填写车号。"),
            (r"^Driver workflow Excel rows: (.+); blank car numbers: (.+)\.$", "车号流程 Excel 行数：{0}；空白车号：{1}。"),
            (r"^Scanning (.+) waybill images\.$", "正在扫描 {0} 张运单图片。"),
            (r"^Driver numbers filled: (.+)\. Already filled: (.+)\.$", "已填写车号：{0}。已存在车号：{1}。"),
            (r"^Container OCR exact matches: (.+); close container matches: (.+)\.$", "箱号 OCR 精确匹配：{0}；安全近似匹配：{1}。"),
            (r"^Review required: (.+) image\(s\) were not filled automatically\.$", "需要复查：{0} 张图片未自动填写。"),
            (r"^Driver workbook: (.+)$", "车号工作簿：{0}"),
            (r"^Output: (.+)$", "输出：{0}"),
            (r"^Bulk PDF: (.+)$", "合并 PDF：{0}"),
            (r"^Creating bulk PDF: (\d+)% \((\d+)/(\d+)\)$", "正在生成合并 PDF：{0}%（{1}/{2}）"),
            (r"^Unlocking files: (\d+)% \((\d+)/(\d+)\)$", "正在解锁文件：{0}%（{1}/{2}）"),
            (r"^Loaded (\d+) container rows from Excel\.$", "已从 Excel 读取 {0} 个箱号。"),
            (r"^Scanning (\d+) PDF files\.$", "正在扫描 {0} 个 PDF 文件。"),
            (r"^Scanning (\d+) PDF and Excel files for removable protection\.$", "正在检查 {0} 个 PDF 和 Excel 文件是否可解锁。"),
            (r"^Finished\. Sorted files are in: (.+)$", "完成。排序文件夹：{0}"),
            (r"^Check sort_report\.csv for missing or unmatched files\.$", "请查看 sort_report.csv 中缺少或未匹配的文件。"),
            (r"^Creating bulk PDF from sorted waybills\.$", "正在按排序结果生成合并 PDF。"),
            (r"^Bulk PDF created: (.+)$", "合并 PDF 已生成：{0}"),
            (r"^Copy verification: (\d+)/(\d+) sorted PDFs are byte-for-byte identical to the downloads\.$", "复制校验：{0}/{1} 个排序 PDF 与下载文件逐字节一致。"),
            (r"^Safe bulk PDF mode: original PDF pages only; no separator pages or digital sequence marks\.$", "安全合并 PDF 模式：只使用原始 PDF 页面，不加分隔页或电子序号。"),
            (r"^Optional bulk PDF extras enabled: (.+)\.$", "已开启可选合并 PDF 功能：{0}。"),
            (
                r"^File unlock finished: (\d+) unlocked; (\d+) already unlocked; (\d+) skipped; (\d+) failed\.$",
                "文件解锁完成：已解锁 {0}；已可用 {1}；已跳过 {2}；失败 {3}。",
            ),
        )
        for pattern, replacement in patterns:
            match = re.match(pattern, message)
            if match:
                return replacement.format(*match.groups())
        match = re.match(r"^(\[\d+/\d+\]) (.+) -> not matched$", message)
        if match:
            return f"{match.group(1)} {match.group(2)} -> 未匹配"
        match = re.match(r"^(\[\d+/\d+\]) (.+) -> ([A-Z]{4}\d{7}) \((.+)\)$", message)
        if match:
            return f"{match.group(1)} {match.group(2)} -> {match.group(3)}"
        match = re.match(r"^(\[\d+/\d+\]) (.+) -> unlocked$", message)
        if match:
            return f"{match.group(1)} {match.group(2)} -> 已解锁"
        match = re.match(r"^(\[\d+/\d+\]) (.+) -> already unlocked$", message)
        if match:
            return f"{match.group(1)} {match.group(2)} -> 已可用"
        match = re.match(r"^(\[\d+/\d+\]) (.+) -> skipped; open password required$", message)
        if match:
            return f"{match.group(1)} {match.group(2)} -> 已跳过；需要打开密码"
        match = re.match(r"^(\[\d+/\d+\]) (.+) -> failed: (.+)$", message)
        if match:
            return f"{match.group(1)} {match.group(2)} -> 失败：{match.group(3)}"
        return message

    def clear_log(self) -> None:
        self.log_history.clear()
        if hasattr(self, "log_text"):
            self.log_text.delete("1.0", tk.END)

    def toggle_language(self) -> None:
        old_language = self.language.get()
        self.language.set("zh" if self.language.get() == "en" else "en")
        self.title(self.tr("window_title"))
        self.status_text.set(self.tr(self.status_key))
        self._refresh_current_screen(old_language=old_language)

    def toggle_theme(self) -> None:
        """Switch between the light and dark colour schemes.

        The setting is saved immediately so it survives a restart. Existing
        widgets are recoloured in place, keeping the window continuously open.
        """
        new_scheme = LIGHT if theme.active() == DARK else DARK
        old_tokens = _load_colour_tokens()
        self.user_settings = app_settings.update_settings(theme=new_scheme)
        apply_theme(new_scheme)
        self.configure(bg=BG_WINDOW)
        self.status_text.set(self.tr(self.status_key))
        self._refresh_current_screen(old_tokens=old_tokens)

    def _refresh_current_screen(
        self,
        old_language: str | None = None,
        old_tokens: dict[str, str] | None = None,
    ) -> None:
        """Refresh existing widgets in place after a language or theme change.

        Keeping the existing widget tree avoids a close/reopen effect, preserves
        file selections and dialog state, and leaves every control in place.
        """
        geometry = self._window_geometry()
        old_language = old_language or self.language.get()
        if old_tokens is not None:
            new_tokens = _load_colour_tokens()
            colour_map: dict[tuple[int, int, int], dict[str, str]] = {}
            for name, previous in old_tokens.items():
                current = new_tokens.get(name)
                if current is None or previous == current:
                    continue
                try:
                    # Tk normalizes some colors returned by cget (case and
                    # channel precision vary by widget/platform). Keep every
                    # token for an RGB value: palettes intentionally reuse
                    # colors (for example SURFACE and ON_ACCENT are both
                    # white in light mode), but those roles may map to
                    # different colors in the other scheme.
                    colour_map.setdefault(self.winfo_rgb(previous), {})[name] = current
                except tk.TclError:
                    continue
        else:
            colour_map = {}

        exact_text = {
            str(previous): str(current)
            for key, previous in I18N[old_language].items()
            if (current := I18N[self.language.get()].get(key)) is not None
        }
        old_messages = I18N[old_language]
        new_messages = I18N[self.language.get()]
        colour_options = (
            "background", "foreground", "activebackground", "activeforeground",
            "disabledforeground", "highlightbackground", "highlightcolor",
            "insertbackground", "readonlybackground", "selectbackground",
            "selectforeground", "troughcolor",
        )
        colour_roles = {
            "background": {
                "BG_WINDOW", "BG_PANEL", "BG_LIST", "BG_WARNING", "SURFACE",
                "ACCENT", "ACCENT_HOVER", "ACCENT_DISABLED", "ACCENT_SOFT",
                "ACCENT_SOFT_HOVER", "BORDER", "DIVIDER", "BORDER_WARNING",
                "SCROLLBAR", "SCROLLBAR_TROUGH", "STATE_SUCCESS", "STATE_ERROR",
                "STATE_WARNING",
            },
            "foreground": {
                "ON_ACCENT", "ACCENT_TEXT", "TEXT", "TEXT_STRONG", "TEXT_HEADING",
                "TEXT_MUTED", "TEXT_FAINT", "TEXT_LABEL", "TEXT_ENTRY",
                "STATE_SUCCESS", "STATE_ERROR", "STATE_WARNING",
            },
            "activebackground": {"ACCENT_HOVER", "ACCENT_SOFT_HOVER", "ACCENT_DISABLED"},
            "activeforeground": {"ON_ACCENT", "TEXT_HEADING"},
            "disabledforeground": {"ON_ACCENT"},
            "highlightbackground": {"BG_WINDOW", "BORDER", "BORDER_WARNING"},
            "highlightcolor": {"ACCENT"},
            "insertbackground": {"TEXT_ENTRY"},
            "readonlybackground": {"BG_LIST"},
            "selectbackground": {"ACCENT", "ACCENT_SOFT"},
            "selectforeground": {"TEXT", "ON_ACCENT"},
            "troughcolor": {"SCROLLBAR_TROUGH"},
        }

        def template_translation(text: str) -> str | None:
            """Translate a formatted I18N string, including repeated fields."""
            for key, template in old_messages.items():
                if not isinstance(template, str) or "{" not in template:
                    continue
                pieces: list[str] = []
                field_groups: dict[str, str] = {}
                cursor = 0
                for index, field_match in enumerate(re.finditer(r"\{([^{}]+)\}", template)):
                    pieces.append(re.escape(template[cursor:field_match.start()]))
                    field = field_match.group(1)
                    group = field_groups.get(field)
                    if group is None:
                        group = f"value_{len(field_groups)}"
                        field_groups[field] = group
                        pieces.append(f"(?P<{group}>.+?)")
                    else:
                        # Repeated placeholders (for example {matched}/{matched})
                        # must refer to the same captured value. Re-declaring
                        # the same named group raises re.PatternError and used
                        # to abort the entire widget refresh partway through.
                        pieces.append(f"(?P={group})")
                    cursor = field_match.end()
                pieces.append(re.escape(template[cursor:]))
                try:
                    match = re.fullmatch("".join(pieces), text, flags=re.DOTALL)
                except re.error:
                    continue
                if match is None:
                    continue
                values = {field: match.group(group) for field, group in field_groups.items()}
                try:
                    return str(new_messages.get(key, template)).format(**values)
                except (KeyError, IndexError, ValueError):
                    continue
            return None

        def colour_replacement(widget: tk.Misc, option: str, value: str) -> str | None:
            try:
                candidates = colour_map.get(self.winfo_rgb(value), {})
            except tk.TclError:
                return None
            role_names = colour_roles.get(option, set())
            matches = {name: replacement for name, replacement in candidates.items() if name in role_names}
            if not matches:
                return None
            # In dark mode TEXT_HEADING and TEXT_ENTRY intentionally share an
            # RGB value. Entries still need their input-text color; all other
            # widgets using that value are headings.
            if option == "foreground" and widget.winfo_class() == "Entry" and "TEXT_ENTRY" in matches:
                return matches["TEXT_ENTRY"]
            preferred = next(iter(matches))
            return matches[preferred]

        def translate_text(text: str) -> str | None:
            translated = exact_text.get(text)
            if translated is None and text.startswith("\u2190  "):
                translated = "\u2190  " + exact_text.get(text[3:], text[3:])
            if translated is None:
                # Translate strings containing values substituted into a
                # localized template, such as a page-count summary.
                translated = template_translation(text)
            return translated

        def refresh_tree(widget: tk.Misc) -> None:
            try:
                children = widget.winfo_children()
            except tk.TclError:
                return
            # Recurse independently of a widget's supported options: one
            # unusual native/ttk option must not prevent its descendants from
            # receiving their text or theme refresh.
            for child in children:
                refresh_tree(child)
            try:
                keys = set(widget.keys())
            except tk.TclError:
                return
            changes: dict[str, str] = {}
            try:
                if "text" in keys:
                    text = str(widget.cget("text"))
                    translated = translate_text(text)
                    if translated is not None and translated != text:
                        changes["text"] = translated
                if "textvariable" in keys and widget.winfo_class() in {"Label", "TLabel", "Message"}:
                    variable_name = str(widget.cget("textvariable"))
                    if variable_name:
                        try:
                            value = str(self.getvar(variable_name))
                            translated = translate_text(value)
                            if translated is not None and translated != value:
                                self.setvar(variable_name, translated)
                        except tk.TclError:
                            pass
                for option in colour_options:
                    if option in keys:
                        value = str(widget.cget(option))
                        replacement = colour_replacement(widget, option, value)
                        if replacement:
                            changes[option] = replacement
                if changes:
                    try:
                        widget.configure(**changes)
                    except tk.TclError:
                        # Apply supported options individually if one widget
                        # rejects a combined option update.
                        for option, value in changes.items():
                            try:
                                widget.configure(**{option: value})
                            except tk.TclError:
                                pass
            except (tk.TclError, TypeError):
                pass

        self._configure_base_style()
        self.title(self.tr("window_title"))
        for child in self.winfo_children():
            refresh_tree(child)
        for attribute, title_key in (("settings_window", "settings_title"), ("options_window", "options_title")):
            dialog = getattr(self, attribute, None)
            try:
                if dialog is not None and dialog.winfo_exists():
                    dialog.title(self.tr(title_key))
                    refresh_tree(dialog)
            except tk.TclError:
                pass
        # Tk's themed button is recreated by neither operation; update its
        # switching label explicitly, while the stable header shell fixes its
        # geometry in both languages.
        theme_button = getattr(self, "theme_button", None)
        if theme_button is not None:
            try:
                label = self.tr("theme_to_light") if theme.active() == DARK else self.tr("theme_to_dark")
                icon = "\u2600" if theme.active() == DARK else "\u263d"
                theme_button.configure(text=f"{icon}  {label}")
            except tk.TclError:
                pass
        if (old_tokens is not None or old_language != self.language.get()) and hasattr(self, "log_text"):
            try:
                self.render_log_history()
            except tk.TclError:
                pass
        self.update_idletasks()
        self._restore_window_geometry(geometry)

    def _window_geometry(self) -> str:
        """Capture the current window geometry.

        Returns the literal 'zoomed' when the window is maximised, otherwise a
        ``WxH+X+Y`` string. While zoomed, reapplying an explicit size would
        un-maximise the window, so the two cases are kept distinct.
        """
        try:
            if self.state() == "zoomed":
                return "zoomed"
            return str(self.wm_geometry())
        except tk.TclError:
            return ""

    def _restore_window_geometry(self, saved: str) -> None:
        """Reapply geometry captured by ``_window_geometry``.

        Applied immediately and then once more on the next idle cycle so text
        layout recalculation cannot change the outer window size.
        """
        if not saved:
            return

        def apply() -> None:
            try:
                if saved == "zoomed":
                    self.state("zoomed")
                else:
                    self.geometry(saved)
            except tk.TclError:
                pass

        apply()
        try:
            self.after_idle(apply)
        except tk.TclError:
            pass

    def set_status(self, status_key: str) -> None:
        self.status_key = status_key
        self.status_text.set(self.tr(status_key))

    def set_busy(self, is_busy: bool, status_key: str | None = None) -> None:
        self.worker_running = is_busy
        if status_key:
            self.set_status(status_key)
        self.update_action_state()

    def update_action_state(self) -> None:
        can_act = not self.worker_running
        if self._button_exists("sort_button"):
            has_result = self.latest_result is not None
            self._set_primary_button_enabled(self.sort_button, can_act)
            self._set_primary_button_enabled(self.make_pdf_button, can_act and has_result)
            self._set_primary_button_enabled(self.open_output_button, can_act and has_result)
        if self._button_exists("driver_fill_button"):
            ready = bool(self.driver_excel_path.get() and self.driver_image_folder.get())
            self._set_primary_button_enabled(self.driver_fill_button, can_act and ready)
            self._set_primary_button_enabled(self.driver_open_output_button, can_act and self.latest_driver_result is not None)
        if self._button_exists("unlock_button"):
            ready = bool(self.unlock_file_paths)
            self._set_primary_button_enabled(self.unlock_button, can_act and ready)
            self._set_primary_button_enabled(
                self.unlock_open_output_button,
                can_act and self.latest_unlock_result is not None,
            )

    def choose_driver_excel(self) -> None:
        if self.worker_running:
            return
        selected = filedialog.askopenfilename(
            title=self.tr("choose_excel"),
            filetypes=(("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")),
        )
        if not selected:
            return
        path = Path(selected)
        try:
            rows, _header_row, blank_count, filled_count = load_driver_excel_rows(path)
        except Exception as exc:
            messagebox.showerror(self.tr("excel_problem"), str(exc))
            self.set_status("status_driver_start")
            return

        self.driver_excel_path.set(str(path))
        self.driver_row_count.set(str(len(rows)))
        self.driver_blank_count.set(str(blank_count))
        self.latest_driver_result = None
        self.set_status("status_driver_excel_loaded" if not self.driver_image_folder.get() else "status_driver_ready")
        self.append_log(f"Driver Excel loaded: {path.name}")
        self.append_log(f"Driver workflow rows: {len(rows)}; blank car numbers: {blank_count}")
        if filled_count:
            self.append_log(f"Rows already containing car numbers: {filled_count}")
        self.update_action_state()

    def choose_driver_image_folder(self) -> None:
        if self.worker_running:
            return
        initial_dir = self.driver_image_folder.get() or (
            str(Path(self.driver_excel_path.get()).parent) if self.driver_excel_path.get() else ""
        )
        dialog_options = {
            "title": self.tr("choose_image_in_folder"),
            "filetypes": (
                ("Image files", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff"),
                ("All files", "*.*"),
            ),
        }
        if initial_dir:
            dialog_options["initialdir"] = initial_dir
        selected = filedialog.askopenfilename(**dialog_options)
        if not selected:
            return
        folder = Path(selected).parent
        images = list_scan_images(folder)
        self.driver_image_folder.set(str(folder))
        self.driver_image_count.set(str(len(images)))
        self.latest_driver_result = None
        self.set_status("status_driver_ready" if self.driver_excel_path.get() else "status_driver_start")
        self.append_log(f"Image folder loaded: {folder}")
        self.append_log(f"Image files: {len(images)}")
        self.update_action_state()

    def validate_driver_inputs(self) -> tuple[Path, Path] | None:
        if not self.driver_excel_path.get():
            messagebox.showwarning(self.tr("missing_excel_title"), self.tr("missing_excel"))
            return None
        if not self.driver_image_folder.get():
            messagebox.showwarning(self.tr("missing_images_title"), self.tr("missing_images"))
            return None
        return Path(self.driver_excel_path.get()), Path(self.driver_image_folder.get())

    def start_driver_fill(self) -> None:
        paths = self.validate_driver_inputs()
        if paths is None or self.worker_running:
            return

        excel_path, image_folder = paths
        self.set_busy(True, "status_driver_working")
        self.append_log("Starting driver number fill.")

        def worker() -> None:
            try:
                result = fill_driver_numbers_from_images(
                    excel_path,
                    image_folder,
                    log=lambda message: self.worker_queue.put(("log", message)),
                )
            except Exception as exc:
                self.worker_queue.put(("driver_error", exc))
            else:
                self.worker_queue.put(("driver_done", result))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self.process_worker_queue)

    def open_driver_output(self) -> None:
        if self.latest_driver_result is None:
            return
        try:
            import os

            os.startfile(self.latest_driver_result.output_dir)
        except Exception as exc:
            messagebox.showerror(self.tr("driver_fill_failed"), str(exc))

    def _set_unlock_sources(self, sources: Iterable[Path], label: str) -> None:
        self.unlock_file_paths = tuple(Path(path) for path in sources)
        self.unlock_source_label.set(label)
        file_count = 0
        for path in self.unlock_file_paths:
            if path.is_file() and path.suffix.lower() in UNLOCKABLE_EXTENSIONS:
                file_count += 1
            elif path.is_dir():
                file_count += sum(
                    1 for item in path.rglob("*")
                    if item.is_file() and item.suffix.lower() in UNLOCKABLE_EXTENSIONS
                    and not any(part.startswith("unlocked_files_") for part in item.relative_to(path).parts)
                )
        self.unlock_pdf_count.set(str(file_count))
        self.latest_unlock_result = None
        self.set_status("status_unlock_ready")
        self.append_log(f"Unlock input loaded: {label}")
        self.append_log(f"Selected paths: {len(self.unlock_file_paths)}")
        self.update_action_state()

    def choose_unlock_files(self) -> None:
        if self.worker_running:
            return
        initial_dir = str(self.unlock_file_paths[0].parent) if self.unlock_file_paths else ""
        dialog_options = {
            "title": self.tr("choose_unlock_files"),
            "filetypes": [
                ("Supported files", "*.pdf *.xlsx *.xlsm *.xltx *.xltm"),
                ("PDF files", "*.pdf"),
                ("Excel files", "*.xlsx *.xlsm *.xltx *.xltm"),
                ("All files", "*.*"),
            ],
        }
        if initial_dir:
            dialog_options["initialdir"] = initial_dir
        selected = filedialog.askopenfilenames(**dialog_options)
        if not selected:
            return
        paths = tuple(Path(path) for path in selected if Path(path).suffix.lower() in UNLOCKABLE_EXTENSIONS)
        self._set_unlock_sources(paths, f"{len(paths)} file(s) selected")

    def choose_unlock_folder(self) -> None:
        if self.worker_running:
            return
        initial_dir = str(self.unlock_file_paths[0]) if self.unlock_file_paths and self.unlock_file_paths[0].is_dir() else ""
        dialog_options = {"title": self.tr("choose_unlock_folder")}
        if initial_dir:
            dialog_options["initialdir"] = initial_dir
        selected = filedialog.askdirectory(**dialog_options)
        if not selected:
            return
        folder = Path(selected)
        self._set_unlock_sources((folder,), str(folder))

    def validate_unlock_inputs(self) -> tuple[Path, ...] | None:
        if not self.unlock_file_paths:
            messagebox.showwarning(self.tr("missing_folder_title"), self.tr("missing_unlock_folder"))
            return None
        return self.unlock_file_paths

    def start_unlock_pdfs(self) -> None:
        sources = self.validate_unlock_inputs()
        if sources is None or self.worker_running:
            return

        self.set_busy(True, "status_unlock_working")
        self.set_progress(0, "0%")
        self.append_log("Starting file unlock.")
        password = self.unlock_password.get()

        def worker() -> None:
            try:
                result = unlock_files(
                    sources,
                    password=password,
                    log=lambda message: self.worker_queue.put(("log", message)),
                    progress=lambda current, total: self.worker_queue.put(("unlock_progress", (current, total))),
                )
            except Exception as exc:
                self.worker_queue.put(("unlock_error", exc))
            else:
                self.worker_queue.put(("unlock_done", result))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self.process_worker_queue)

    def open_unlock_output(self) -> None:
        if self.latest_unlock_result is None:
            return
        try:
            import os

            os.startfile(self.latest_unlock_result.output_dir)
        except Exception as exc:
            messagebox.showerror(self.tr("unlock_failed"), str(exc))

    def load_sort_excel(
        self,
        path: Path,
        manual_columns: tuple[int, int, int] | None = None,
    ) -> None:
        amount, rule = lookup_excel_copy_amount(path)
        if manual_columns:
            header_row, sequence_column, container_column = manual_columns
            rows = load_waybill_rows(path, header_row, sequence_column, container_column)
        else:
            rows = load_waybill_rows(path)

        self.excel_column_override = manual_columns
        self.excel_path.set(str(path))
        self.copy_amount.set(f"{amount} ({rule})")
        self.container_count.set(str(len(rows)))
        self.latest_result = None
        self.latest_print_batch = None
        self.latest_print_signature = None
        self.set_status("status_excel_loaded")
        self.append_log(f"Excel loaded: {path.name}")
        self.append_log(f"Excel supplementary count: {amount} ({rule})")
        if manual_columns:
            header_row, sequence_column, container_column = manual_columns
            self.append_log(
                "Excel columns: "
                f"header row {header_row}; "
                f"sequence {excel_column_name(sequence_column)}; "
                f"container {excel_column_name(container_column)}"
            )
        self.append_log(f"Container rows: {len(rows)}")
        if self.pdf_folder.get():
            self.update_pdf_folder_analysis(Path(self.pdf_folder.get()))
        self.update_action_state()

    def default_excel_column_setup(self, path: Path) -> tuple[int, int, int]:
        if self.excel_column_override:
            return self.excel_column_override
        try:
            workbook = load_workbook(path, data_only=True, read_only=True)
            worksheet = workbook.active
            header_row, columns = find_header_columns(worksheet, (SEQ_HEADER, CONTAINER_HEADER))
            return header_row, columns[SEQ_HEADER], columns[CONTAINER_HEADER]
        except Exception:
            return 1, 1, 3

    def open_excel_columns_dialog(
        self,
        path: Path | None = None,
        initial_error: str | None = None,
    ) -> None:
        if self.worker_running:
            return
        if path is None:
            if not self.excel_path.get():
                messagebox.showinfo(self.tr("missing_excel_title"), self.tr("missing_excel"))
                return
            path = Path(self.excel_path.get())

        header_row, sequence_column, container_column = self.default_excel_column_setup(path)
        window = tk.Toplevel(self)
        window.title(self.tr("excel_columns_title"))
        window.configure(bg=SURFACE)
        window.transient(self)
        window.grab_set()

        body = tk.Frame(window, bg=SURFACE, padx=20, pady=18)
        body.pack(fill=tk.BOTH, expand=True)
        tk.Label(
            body,
            text=self.tr("excel_columns_note"),
            bg=SURFACE,
            fg=TEXT_MUTED,
            font=("Segoe UI", 10),
            justify="left",
            wraplength=420,
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        if initial_error:
            tk.Label(
                body,
                text=initial_error,
                bg=BG_WARNING,
                fg=STATE_WARNING,
                font=("Segoe UI", 9),
                justify="left",
                wraplength=420,
                padx=10,
                pady=8,
            ).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 12))

        header_var = tk.StringVar(value=str(header_row))
        sequence_var = tk.StringVar(value=excel_column_name(sequence_column))
        container_var = tk.StringVar(value=excel_column_name(container_column))

        fields = (
            (self.tr("header_row"), header_var),
            (self.tr("sequence_column"), sequence_var),
            (self.tr("container_column"), container_var),
        )
        start_row = 2
        for row_index, (label, variable) in enumerate(fields, start=start_row):
            tk.Label(body, text=label, bg=SURFACE, fg=TEXT_HEADING, font=("Segoe UI", 10, "bold")).grid(
                row=row_index,
                column=0,
                sticky="w",
                pady=6,
                padx=(0, 12),
            )
            tk.Entry(body, textvariable=variable, font=("Segoe UI", 11), width=12).grid(
                row=row_index,
                column=1,
                sticky="ew",
                pady=6,
            )
        body.columnconfigure(1, weight=1)

        buttons = tk.Frame(body, bg=SURFACE)
        buttons.grid(row=start_row + len(fields), column=0, columnspan=2, sticky="e", pady=(16, 0))

        def apply_manual_columns() -> None:
            parsed_header = parse_sequence(header_var.get())
            parsed_sequence = parse_excel_column(sequence_var.get())
            parsed_container = parse_excel_column(container_var.get())
            if parsed_header is None or parsed_sequence is None or parsed_container is None:
                messagebox.showerror(self.tr("excel_problem"), "Enter a header row number and valid Excel columns, such as A and C.")
                return
            try:
                self.load_sort_excel(path, (parsed_header, parsed_sequence, parsed_container))
            except Exception as exc:
                messagebox.showerror(self.tr("excel_problem"), str(exc))
                return
            window.destroy()

        self._secondary_button(
            buttons, self.tr("cancel"), window.destroy, label_key="cancel"
        ).pack(side=tk.RIGHT, padx=(8, 0))
        self._primary_button(buttons, self.tr("apply"), apply_manual_columns, compact=True).pack(side=tk.RIGHT)

        window.update_idletasks()
        x = self.winfo_rootx() + max(40, (self.winfo_width() - window.winfo_width()) // 2)
        y = self.winfo_rooty() + max(40, (self.winfo_height() - window.winfo_height()) // 3)
        window.geometry(f"+{x}+{y}")

    def choose_excel(self) -> None:
        if self.worker_running:
            return
        selected = filedialog.askopenfilename(
            title=self.tr("choose_excel"),
            filetypes=(("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")),
        )
        if not selected:
            return
        path = Path(selected)
        amount, rule = lookup_excel_copy_amount(path)
        try:
            self.load_sort_excel(path)
        except Exception as exc:
            message = f"Excel supplementary count: {amount} ({rule})\n\n{exc}"
            if "Could not find required Excel columns" in str(exc):
                if messagebox.askyesno(self.tr("excel_problem"), message + "\n\nUse manual Excel column setup?"):
                    self.open_excel_columns_dialog(path, str(exc))
            else:
                messagebox.showerror(self.tr("excel_problem"), message)
            self.set_status("status_start")
            return

    def choose_pdf_folder(self) -> None:
        if self.worker_running:
            return
        initial_dir = self.pdf_folder.get() or (str(Path(self.excel_path.get()).parent) if self.excel_path.get() else "")
        dialog_options = {
            "title": self.tr("choose_pdf_in_folder"),
            "filetypes": (("PDF files", "*.pdf"), ("All files", "*.*")),
        }
        if initial_dir:
            dialog_options["initialdir"] = initial_dir
        selected = filedialog.askopenfilename(**dialog_options)
        if not selected:
            return
        path = Path(selected).parent
        self.pdf_folder.set(str(path))
        self.latest_result = None
        self.latest_print_batch = None
        self.latest_print_signature = None
        self.set_status("status_ready")
        self.append_log(f"PDF folder loaded: {path}")
        self.update_pdf_folder_analysis(path)
        self.update_action_state()

    def validate_inputs(self) -> tuple[Path, Path] | None:
        if not self.excel_path.get():
            messagebox.showwarning(self.tr("missing_excel_title"), self.tr("missing_excel"))
            return None
        if not self.pdf_folder.get():
            messagebox.showwarning(self.tr("missing_folder_title"), self.tr("missing_folder"))
            return None
        return Path(self.excel_path.get()), Path(self.pdf_folder.get())

    def start_sort(self) -> None:
        paths = self.validate_inputs()
        if paths is None or self.worker_running:
            return

        excel_path, pdf_folder = paths
        self.set_busy(True, "status_sorting")
        self.append_log("Starting sort.")

        def worker() -> None:
            try:
                result = sort_waybills(
                    excel_path,
                    pdf_folder,
                    log=lambda message: self.worker_queue.put(("log", message)),
                )
            except Exception as exc:
                self.worker_queue.put(("error", exc))
            else:
                self.worker_queue.put(("done", result))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self.process_worker_queue)

    def require_sort_result(self) -> SortResult | None:
        if self.latest_result is None:
            messagebox.showwarning(self.tr("missing_sort_title"), self.tr("missing_sort"))
            return None
        return self.latest_result

    def current_print_signature(self, result: SortResult) -> tuple[str, bool, bool]:
        return (str(result.report_path), bool(self.include_separators.get()), bool(self.mark_sequence.get()))

    def create_print_pdf(self) -> None:
        result = self.require_sort_result()
        if result is None or self.worker_running:
            return
        separators = bool(self.include_separators.get())
        mark_sequence = bool(self.mark_sequence.get())
        signature = self.current_print_signature(result)

        if self.latest_print_batch and self.latest_print_signature == signature and self.latest_print_batch.path.exists():
            messagebox.showinfo(
                self.tr("pdf_finished"),
                self.tr(
                    "pdf_finished_msg",
                    path=self.latest_print_batch.path,
                    pages=self.latest_print_batch.output_pages,
                    separators=self.latest_print_batch.separator_pages,
                ),
            )
            return

        if (separators or mark_sequence) and not messagebox.askyesno(
            self.tr("optional_pdf_features_warning_title"),
            self.tr("optional_pdf_features_warning"),
        ):
            return

        self.set_busy(True, "status_creating_pdf")
        self.set_progress(0, "0%")

        def worker() -> None:
            try:
                batch = create_print_batch_pdf(
                    result.report_path,
                    result.output_dir,
                    separators,
                    mark_sequence,
                    log=lambda message: self.worker_queue.put(("log", message)),
                    progress=lambda current, total: self.worker_queue.put(("progress", (current, total))),
                )
                self.worker_queue.put(("pdf_done", batch))
            except Exception as exc:
                self.worker_queue.put(("pdf_error", exc))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self.process_worker_queue)

    def open_output(self) -> None:
        result = self.require_sort_result()
        if result is None:
            return
        try:
            import os

            os.startfile(result.output_dir)
        except Exception as exc:
            messagebox.showerror(self.tr("sort_failed"), str(exc))

    def process_worker_queue(self) -> None:
        try:
            while True:
                kind, payload = self.worker_queue.get_nowait()
                if kind == "log":
                    self.append_log(str(payload))
                elif kind == "progress":
                    current, total = payload
                    percent = int((current / total) * 100) if total else 0
                    text = f"{percent}% ({current}/{total})"
                    self.set_progress(percent, text)
                    if current == total or current == 1 or current % 5 == 0:
                        self.append_log(f"Creating bulk PDF: {percent}% ({current}/{total})")
                elif kind == "unlock_progress":
                    current, total = payload
                    percent = int((current / total) * 100) if total else 0
                    text = f"{percent}% ({current}/{total})"
                    self.set_progress(percent, text)
                    if current == total or current == 1 or current % 5 == 0:
                        self.append_log(f"Unlocking files: {percent}% ({current}/{total})")
                elif kind == "error":
                    self.set_busy(False)
                    self.status_text.set(self.tr("sort_failed"))
                    self.append_log(f"ERROR: {payload}")
                    messagebox.showerror(self.tr("sort_failed"), str(payload))
                elif kind == "pdf_error":
                    self.set_busy(False)
                    self.status_text.set(self.tr("pdf_failed"))
                    self.append_log(f"ERROR: {payload}")
                    messagebox.showerror(self.tr("pdf_failed"), str(payload))
                elif kind == "driver_error":
                    self.set_busy(False)
                    self.status_text.set(self.tr("driver_fill_failed"))
                    self.append_log(f"ERROR: {payload}")
                    messagebox.showerror(self.tr("driver_fill_failed"), str(payload))
                elif kind == "unlock_error":
                    self.set_busy(False)
                    self.status_text.set(self.tr("unlock_failed"))
                    self.append_log(f"ERROR: {payload}")
                    messagebox.showerror(self.tr("unlock_failed"), str(payload))
                elif kind == "done":
                    result = payload
                    assert isinstance(result, SortResult)
                    self.latest_result = result
                    self.latest_print_batch = None
                    self.latest_print_signature = None
                    self.set_busy(False, "status_finished")
                    self.append_log(f"Output: {result.output_dir}")
                    messagebox.showinfo(
                        self.tr("sort_finished"),
                        self.tr(
                            "sort_finished_msg",
                            sorted_dir=result.sorted_dir,
                            matched=result.matched_count,
                            verified=result.verified_copy_count,
                            missing=result.missing_count,
                            unmatched=result.unmatched_count,
                        ),
                    )
                elif kind == "unlock_done":
                    result = payload
                    assert isinstance(result, PdfUnlockResult)
                    self.latest_unlock_result = result
                    self.set_busy(False, "status_unlock_ready")
                    self.set_progress(100, "100%")
                    self.append_log(f"Output: {result.output_dir}")
                    messagebox.showinfo(
                        self.tr("unlock_finished"),
                        self.tr(
                            "unlock_finished_msg",
                            output_dir=result.output_dir,
                            unlocked=result.unlocked_count,
                            already=result.already_unlocked_count,
                            skipped=result.skipped_count,
                            failed=result.failed_count,
                            report=result.report_path,
                        ),
                    )
                elif kind == "driver_done":
                    result = payload
                    assert isinstance(result, DriverFillResult)
                    self.latest_driver_result = result
                    self.set_busy(False, "status_finished")
                    self.append_log(f"Output: {result.output_dir}")
                    messagebox.showinfo(
                        self.tr("driver_fill_finished"),
                        self.tr(
                            "driver_fill_finished_msg",
                            workbook=result.workbook_path,
                            filled=result.filled_count,
                            already=result.already_filled_count,
                            review=result.review_count,
                            unmatched=result.unmatched_count,
                        ),
                    )
                elif kind == "pdf_done":
                    batch = payload
                    assert isinstance(batch, PrintBatchResult)
                    self.latest_print_batch = batch
                    if self.latest_result:
                        self.latest_print_signature = self.current_print_signature(self.latest_result)
                    self.set_busy(False, "status_finished")
                    self.set_progress(100, "100%")
                    self.append_log(f"Bulk PDF: {batch.path}")
                    messagebox.showinfo(
                        self.tr("pdf_finished"),
                        self.tr(
                            "pdf_finished_msg",
                            path=batch.path,
                            pages=batch.output_pages,
                            separators=batch.separator_pages,
                        ),
                    )
        except queue.Empty:
            pass

        if self.worker_running:
            self.after(100, self.process_worker_queue)


def run_cli(args: argparse.Namespace) -> int:
    result = sort_waybills(
        Path(args.excel),
        Path(args.pdf_folder),
        Path(args.output) if args.output else None,
        log=print,
    )
    print(f"Output: {result.output_dir}")
    print(f"Sorted: {result.sorted_dir}")
    print(f"Report: {result.report_path}")
    if args.make_print_pdf:
        batch = create_print_batch_pdf(
            result.report_path,
            result.output_dir,
            args.separators,
            args.mark_sequence,
            log=print,
        )
        print(f"Bulk PDF: {batch.path}")
        print(f"Bulk PDF pages: {batch.output_pages}")
    return 0 if result.missing_count == 0 and result.unmatched_count == 0 else 2


def run_unlock_cli(args: argparse.Namespace) -> int:
    unlock_source = args.unlock_source or args.pdf_folder
    result = unlock_files(
        (Path(unlock_source),),
        password=args.password or "",
        output_root=Path(args.output) if args.output else None,
        log=print,
    )
    print(f"Output: {result.output_dir}")
    print(f"Unlocked PDFs: {result.unlocked_dir}")
    print(f"Report: {result.report_path}")
    return 0 if result.skipped_count == 0 and result.failed_count == 0 else 2


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Sort UniTrans PDF waybills by Excel sequence.")
    parser.add_argument("--excel", help="Path to the Excel workbook.")
    parser.add_argument("--pdf-folder", help="Folder containing PDFs.")
    parser.add_argument("--output", help="Optional folder where the timestamped output folder is created.")
    parser.add_argument("--separators", action="store_true", help="Add separator pages between waybills.")
    parser.add_argument("--make-print-pdf", action="store_true", help="Create one bulk PDF after sorting.")
    parser.add_argument("--mark-sequence", action="store_true", help="Stamp the sequence number on detected form 1 and form 4 pages.")
    parser.add_argument("--unlock-pdfs", action="store_true", help="Unlock supported PDF and Excel files from a file or folder (legacy flag name).")
    parser.add_argument("--unlock-source", help="PDF/Excel file or folder to unlock.")
    parser.add_argument("--password", help="Optional password for encrypted PDF/Excel files.")
    parser.add_argument("--check-updates", action="store_true", help="Check GitHub for a newer release and exit.")
    return parser.parse_args(argv)


def configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def run_check_updates_cli() -> int:
    """``--check-updates``: same check the Settings window performs, no GUI.

    Exit codes: 0 = up to date, 2 = an update is available, 1 = check failed.
    """
    result = check_for_update()
    print(f"Installed version: {result.current_version} ({display_version(result.current_version)})")

    if not result.ok:
        print(f"Update check failed: {result.error}", file=sys.stderr)
        return 1

    if not result.update_available:
        print("You are running the latest version.")
        return 0

    print(f"Update available: {result.remote_version}")
    if result.release_page:
        print(f"Release page: {result.release_page}")
    if result.download_url:
        print(f"Download: {result.download_url}")
    size = format_size(result.size_bytes)
    if size:
        print(f"Size: {size}")
    for note in result.notes_for("en"):
        print(f"  - {note}")
    return 2


def main(argv: list[str] | None = None) -> int:
    configure_stdio()
    args = parse_args(sys.argv[1:] if argv is None else argv)
    if args.check_updates:
        try:
            return run_check_updates_cli()
        except Exception as exc:
            print(f"ERROR: {exc}", file=sys.stderr)
            return 1
    if args.unlock_pdfs:
        if not (args.unlock_source or args.pdf_folder):
            print("--unlock-pdfs requires --unlock-source (or legacy --pdf-folder).", file=sys.stderr)
            return 1
        try:
            return run_unlock_cli(args)
        except Exception as exc:
            print(f"ERROR: {exc}", file=sys.stderr)
            return 1

    if args.excel or args.pdf_folder:
        if not args.excel or not args.pdf_folder:
            print("--excel and --pdf-folder must be used together.", file=sys.stderr)
            return 1
        try:
            return run_cli(args)
        except Exception as exc:
            print(f"ERROR: {exc}", file=sys.stderr)
            return 1

    app = WaybillSorterApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
