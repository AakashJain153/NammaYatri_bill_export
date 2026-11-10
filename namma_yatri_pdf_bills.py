#!/usr/bin/env python3
"""
Namma Yatri ride-invoice organiser.

When you run it, a window opens asking you to pick the folder that contains
the invoice PDFs. Inside that folder it then creates a "Refined" folder with:

  1. A copy of every PDF renamed to
         Namma Yatri bill_YYMMDD_to_YYMMDD.pdf
     (first ride date to last ride date found in that PDF).
     Your original files are left untouched.
  2. "Namma Yatri rides.xlsx" with one row per ride:
         File Name (hyperlinked to the renamed PDF) | Ride ID | Vehicle No |
         Driver Name | Ride Date | Pickup Time | Drop Time | Pickup Address |
         Drop Address | Ride Fare

Install:
    pip install pdfplumber openpyxl pandas

Usage:
    python namma_yatri_bills.py                 # opens the folder picker
    python namma_yatri_bills.py "C:\\path\\to\\folder"   # skip the picker
"""

import re
import shutil
import sys
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import pdfplumber
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

REFINED_DIR = "Refined"
FILE_PREFIX = "Namma Yatri bill"
EXCEL_NAME = "Namma Yatri rides.xlsx"

# ----------------------------------------------------------------------------
# Patterns
# ----------------------------------------------------------------------------
DATE_RE = re.compile(r"^(\d{1,2}\s+[A-Za-z]{3,9},?\s+\d{4})$")          # 28 Aug, 2026
FARE_RE = re.compile(r"Ride\s*Fare\s*:?\s*[₹Rs\.\s]*([\d,]+(?:\.\d+)?)", re.I)
TIME_RE = re.compile(r"^(\d{1,2}:\d{2}\s*[AP]M)$", re.I)                # 9:07 AM
DRIVER_LBL = re.compile(r"^DRIVER\s+NAME$", re.I)
VEHICLE_LBL = re.compile(r"^VEHICLE\s+NUMBER$", re.I)
RIDEID_LBL = re.compile(r"^RIDE\s+ID$", re.I)

COLUMNS = [
    "File Name", "Ride ID", "Vehicle No", "Driver Name", "Ride Date",
    "Pickup Time", "Drop Time", "Pickup Address", "Drop Address", "Ride Fare",
]


# ----------------------------------------------------------------------------
# Folder selection
# ----------------------------------------------------------------------------
def choose_folder() -> Path:
    """CLI argument if given, otherwise a folder-picker window (text prompt fallback)."""
    if len(sys.argv) > 1:
        return Path(sys.argv[1]).expanduser()

    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        chosen = filedialog.askdirectory(title="Select the folder containing the ride PDFs")
        root.destroy()
        if chosen:
            return Path(chosen)
        sys.exit("No folder selected. Exiting.")
    except Exception:
        # No GUI available (e.g. SSH / headless) -> ask in the terminal
        chosen = input("Enter the full path of the folder containing the PDFs: ").strip().strip('"')
        if not chosen:
            sys.exit("No folder given. Exiting.")
        return Path(chosen).expanduser()


# ----------------------------------------------------------------------------
# Parsing
# ----------------------------------------------------------------------------
def pdf_to_lines(pdf_path: Path) -> list[str]:
    """All non-empty text lines from every page, in reading order."""
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines.extend(l.strip() for l in text.splitlines() if l.strip())
    return lines


def parse_date(s: str):
    s = s.replace(",", "")
    for fmt in ("%d %b %Y", "%d %B %Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return s  # fall back to raw text


def parse_lines(lines: list[str]) -> list[dict]:
    """
    A ride starts at a date line (e.g. '28 Aug, 2026') and runs until the next
    date line / end of document. Inside a ride:
        Ride Fare: ₹428
        <time> + following lines  -> pickup (first) / drop (second) address
        DRIVER NAME / VEHICLE NUMBER / RIDE ID -> value on the next line
    Addresses wrapped over several lines are joined.
    """
    rides = []
    ride = None
    stops = []      # [[time, [address lines]], ...]
    mode = None     # what the next plain line means

    def finish():
        nonlocal ride, stops, mode
        if ride is None:
            return
        for idx, key in enumerate(("Pickup", "Drop")):
            if idx < len(stops):
                ride[f"{key} Time"] = stops[idx][0]
                ride[f"{key} Address"] = " ".join(stops[idx][1]).strip()
        rides.append(ride)
        ride, stops, mode = None, [], None

    for line in lines:
        m = DATE_RE.match(line)
        if m:
            finish()
            ride = {c: None for c in COLUMNS}
            ride["Ride Date"] = parse_date(m.group(1))
            stops, mode = [], None
            continue

        if ride is None:            # header text before the first ride
            continue

        m = FARE_RE.search(line)
        if m:
            ride["Ride Fare"] = float(m.group(1).replace(",", ""))
            mode = None
            continue

        if TIME_RE.match(line):
            stops.append([re.sub(r"\s+", " ", line.upper()), []])
            mode = "addr"
            continue

        if DRIVER_LBL.match(line):
            mode = "driver"; continue
        if VEHICLE_LBL.match(line):
            mode = "vehicle"; continue
        if RIDEID_LBL.match(line):
            mode = "rideid"; continue

        if mode == "driver":
            ride["Driver Name"] = line; mode = None
        elif mode == "vehicle":
            ride["Vehicle No"] = line; mode = None
        elif mode == "rideid":
            ride["Ride ID"] = line; mode = None
        elif mode == "addr" and stops:
            stops[-1][1].append(line)

    finish()
    return rides


# ----------------------------------------------------------------------------
# Renaming
# ----------------------------------------------------------------------------
def new_pdf_name(rides: list[dict], used: set[str]) -> str | None:
    """'Namma Yatri bill_YYMMDD_to_YYMMDD.pdf' (adds _2, _3... if a name repeats)."""
    dates = [r["Ride Date"] for r in rides if isinstance(r["Ride Date"], date)]
    if not dates:
        return None
    base = f"{FILE_PREFIX}_{min(dates):%y%m%d}_to_{max(dates):%y%m%d}"
    name, n = f"{base}.pdf", 2
    while name.lower() in used:
        name = f"{base}_{n}.pdf"
        n += 1
    used.add(name.lower())
    return name


# ----------------------------------------------------------------------------
# Excel output
# ----------------------------------------------------------------------------
def write_excel(df: pd.DataFrame, out_path: Path) -> None:
    cols = list(df.columns)
    col_idx = {c: i + 1 for i, c in enumerate(cols)}

    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Rides")
        ws = xw.sheets["Rides"]

        hdr_fill = PatternFill("solid", start_color="1F3864")
        for c in ws[1]:
            c.font = Font(name="Arial", bold=True, color="FFFFFF")
            c.fill = hdr_fill
            c.alignment = Alignment(horizontal="center", vertical="center")

        widths = {"File Name": 42, "Ride ID": 14, "Vehicle No": 14, "Driver Name": 20,
                  "Ride Date": 13, "Pickup Time": 12, "Drop Time": 12,
                  "Pickup Address": 60, "Drop Address": 60, "Ride Fare": 11}
        for c, i in col_idx.items():
            ws.column_dimensions[get_column_letter(i)].width = widths.get(c, 15)

        wrap_cols = {col_idx["Pickup Address"], col_idx["Drop Address"]}
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name="Arial")
                c.alignment = Alignment(vertical="top", wrap_text=c.column in wrap_cols)
            row[col_idx["Ride Fare"] - 1].number_format = '"₹"#,##0.00'
            row[col_idx["Ride Date"] - 1].number_format = "dd-mmm-yyyy"

            # Hyperlink File Name -> renamed PDF (relative link: Excel sits in the same folder)
            fc = row[col_idx["File Name"] - 1]
            if fc.value:
                fc.hyperlink = quote(str(fc.value))
                fc.font = Font(name="Arial", color="0563C1", underline="single")

        last = ws.max_row
        total_row = last + 1
        fl = get_column_letter(col_idx["Ride Fare"])
        ws.cell(total_row, col_idx["Ride Fare"] - 1, "Total").font = Font(name="Arial", bold=True)
        t = ws.cell(total_row, col_idx["Ride Fare"], f"=SUM({fl}2:{fl}{last})")
        t.font = Font(name="Arial", bold=True)
        t.number_format = '"₹"#,##0.00'

        ws.freeze_panes = "B2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{last}"


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def main():
    folder = choose_folder()
    if not folder.is_dir():
        sys.exit(f"Not a folder: {folder}")

    pdfs = sorted(p for p in folder.glob("*.pdf") if p.is_file())
    if not pdfs:
        sys.exit(f"No PDF files found in {folder}")

    refined = folder / REFINED_DIR
    refined.mkdir(exist_ok=True)

    used_names: set[str] = set()
    all_rides: list[dict] = []

    for pdf in pdfs:
        rides = parse_lines(pdf_to_lines(pdf))
        new_name = new_pdf_name(rides, used_names)

        if new_name is None:
            print(f"  ! {pdf.name}: no rides found - copied with its original name")
            shutil.copy2(pdf, refined / pdf.name)
            continue

        shutil.copy2(pdf, refined / new_name)
        print(f"  {pdf.name}  ->  {new_name}  ({len(rides)} ride(s))")
        for r in rides:
            r["File Name"] = new_name
        all_rides.extend(rides)

    if not all_rides:
        sys.exit("No rides could be extracted.")

    df = pd.DataFrame(all_rides, columns=COLUMNS)
    df = df.drop_duplicates(subset="Ride ID", keep="first")      # same invoice supplied twice
    df["_d"] = pd.to_datetime(df["Ride Date"], errors="coerce")
    df["_t"] = pd.to_datetime(df["Pickup Time"], format="%I:%M %p", errors="coerce")
    df = df.sort_values(["_d", "_t"]).drop(columns=["_d", "_t"]).reset_index(drop=True)

    out = refined / EXCEL_NAME
    try:
        write_excel(df, out)
    except PermissionError:
        sys.exit(f"Cannot write {out} - close it in Excel and run again.")

    print(f"\nDone. {len(df)} rides, total fare ₹{df['Ride Fare'].sum():,.2f}")
    print(f"Output folder: {refined.resolve()}")


if __name__ == "__main__":
    main()
