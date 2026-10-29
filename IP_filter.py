#!/usr/bin/env python3
"""Split the "IP" column of an .xlsx file into "Public IP" and "Private IP".

Usage:
    python split_ips.py input.xlsx [-o output.xlsx] [--sheet SHEET] [--header-row N]

- The "IP" column is renamed to "Public IP" and keeps only public addresses.
- A new "Private IP" column is inserted right after it and receives every
  private address (RFC 1918, loopback, link-local, etc.) moved out of "Public IP".
- Values that are not valid IP addresses are left untouched in "Public IP"
  and reported at the end, so nothing is silently lost.
"""
import argparse
import ipaddress
import sys
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def parse_args():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("input", type=Path, help="Input .xlsx file")
    p.add_argument("-o", "--output", type=Path, help="Output file (default: <input>_split.xlsx)")
    p.add_argument("--sheet", help="Sheet name (default: active sheet)")
    p.add_argument("--header-row", type=int, default=1, help="Row containing headers (default: 1)")
    return p.parse_args()


def find_ip_column(ws, header_row):
    for cell in ws[header_row]:
        if isinstance(cell.value, str) and cell.value.strip().lower() == "ip":
            return cell.column
    return None


def classify(value):
    """Return 'private', 'public', or None (empty / not an IP)."""
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        ip = ipaddress.ip_address(text)
    except ValueError:
        return "invalid"
    # is_global is the precise "public internet" test; it excludes private,
    # loopback, link-local, CGNAT (100.64/10), documentation ranges, etc.
    return "public" if ip.is_global else "private"


def main():
    args = parse_args()
    if not args.input.exists():
        sys.exit(f"File not found: {args.input}")
    output = args.output or args.input.with_name(f"{args.input.stem}_split.xlsx")

    wb = load_workbook(args.input)
    ws = wb[args.sheet] if args.sheet else wb.active

    hdr = args.header_row
    ip_col = find_ip_column(ws, hdr)
    if ip_col is None:
        sys.exit(f'No "IP" column found in row {hdr} of sheet "{ws.title}".')

    new_col = ip_col + 1
    ws.insert_cols(new_col)

    # Headers (copy style from the original header cell)
    src_header = ws.cell(row=hdr, column=ip_col)
    new_header = ws.cell(row=hdr, column=new_col)
    src_header.value = "Public IP"
    new_header.value = "Private IP"
    new_header._style = copy(src_header._style)

    # Keep column width consistent
    width = ws.column_dimensions[get_column_letter(ip_col)].width
    if width:
        ws.column_dimensions[get_column_letter(new_col)].width = width

    counts = {"public": 0, "private": 0, "invalid": 0}
    invalid_rows = []

    for row in range(hdr + 1, ws.max_row + 1):
        cell = ws.cell(row=row, column=ip_col)
        kind = classify(cell.value)
        if kind is None:
            continue
        counts[kind] += 1
        if kind == "private":
            target = ws.cell(row=row, column=new_col)
            target.value = str(cell.value).strip()
            target._style = copy(cell._style)
            cell.value = None
        elif kind == "invalid":
            invalid_rows.append((row, cell.value))

    wb.save(output)

    print(f"Sheet: {ws.title}")
    print(f"Public IPs kept : {counts['public']}")
    print(f"Private IPs moved: {counts['private']}")
    print(f"Invalid values   : {counts['invalid']} (left in 'Public IP')")
    for row, val in invalid_rows[:20]:
        print(f"  row {row}: {val!r}")
    if len(invalid_rows) > 20:
        print(f"  ... and {len(invalid_rows) - 20} more")
    print(f"Saved to: {output}")


if __name__ == "__main__":
    main()