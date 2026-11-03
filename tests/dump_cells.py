#!/usr/bin/env python3
"""Test helper: dump cell values / formulas from an SOW Units sheet.

Usage:
  python3 tests/dump_cells.py values   <workbook> <addresses.json>
  python3 tests/dump_cells.py formulas <workbook> <addresses.json>

Prints a JSON object {address: value} on stdout. In values mode the
workbook is read data-only (cached values); in formulas mode a cell
holding a formula reports its formula string, anything else null.
"""
import json
import sys

import openpyxl


def main():
    mode, path, addr_file = sys.argv[1], sys.argv[2], sys.argv[3]
    with open(addr_file) as f:
        addresses = json.load(f)
    wb = openpyxl.load_workbook(path, data_only=(mode == 'values'), read_only=True)
    ws = wb['SOW Units']
    out = {}
    for addr in addresses:
        v = ws[addr].value
        if mode == 'formulas':
            out[addr] = v if isinstance(v, str) and v.startswith('=') else None
        else:
            out[addr] = v
    print(json.dumps(out))


if __name__ == '__main__':
    main()
