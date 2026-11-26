"""Validate, store and report dated inventory snapshots."""
from __future__ import annotations

import sqlite3
from contextlib import closing
from datetime import date
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
DB_PATH = ROOT / "data" / "inventory.db"
HEADERS = ("SKU", "Product", "Brand", "Category", "Area", "Quantity", "Unit Cost")
AREAS = ("AVAILABLE", "RECEIVING", "DAMAGED")

SCHEMA = """
CREATE TABLE IF NOT EXISTS inventory (
    snapshot_date TEXT NOT NULL,
    sku TEXT NOT NULL,
    product TEXT NOT NULL,
    brand TEXT NOT NULL,
    category TEXT NOT NULL,
    area TEXT NOT NULL CHECK (area IN ('AVAILABLE', 'RECEIVING', 'DAMAGED')),
    quantity INTEGER NOT NULL CHECK (quantity >= 0),
    unit_cost_cents INTEGER NOT NULL CHECK (unit_cost_cents >= 0),
    PRIMARY KEY (snapshot_date, sku, area)
);
"""


def connect(db_path=DB_PATH):
    if str(db_path) != ":memory:":
        Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA)
    return conn


def read_snapshot(source):
    """Read the generic Excel template, reject ambiguous or invalid rows."""
    workbook = load_workbook(source, read_only=True, data_only=False)
    try:
        if "Stock" not in workbook.sheetnames:
            raise ValueError("Missing Stock sheet.")
        rows = workbook["Stock"].iter_rows(values_only=True)
        header = next(rows, ())
        if tuple(header) != HEADERS:
            raise ValueError("Invalid headers. Use the supplied template.")
        records, keys, catalog = [], set(), {}
        for number, row in enumerate(rows, 2):
            if all(value is None for value in row):
                continue
            try:
                sku, product, brand, category, area, qty, cost = row
                texts = (sku, product, brand, category, area)
                if not all(isinstance(v, str) and v.strip() and not v.startswith('=') for v in texts):
                    raise ValueError("Required text fields must be nonempty literal text; SKU must be text.")
                sku, product, brand, category, area = (v.strip() for v in texts)
                area = area.upper()
                if area not in AREAS:
                    raise ValueError("Unknown stock area.")
                if isinstance(qty, bool) or isinstance(cost, bool):
                    raise ValueError("Boolean numbers are not allowed.")
                quantity = Decimal(str(qty))
                unit_cost = Decimal(str(cost))
                if not quantity.is_finite() or quantity < 0 or quantity != quantity.to_integral_value():
                    raise ValueError("Quantity must be a nonnegative integer.")
                if not unit_cost.is_finite() or unit_cost < 0 or unit_cost * 100 != (unit_cost * 100).to_integral_value():
                    raise ValueError("Cost must be nonnegative with at most two decimals.")
                cents = int(unit_cost * 100)
                if int(quantity) > 1_000_000_000 or cents > 1_000_000_000 or int(quantity) * cents > 10**15:
                    raise ValueError("Numeric value exceeds supported demo limits.")
                key = (sku, area)
                if key in keys:
                    raise ValueError("Duplicate SKU and area.")
                metadata = (product, brand, category, cents)
                if sku in catalog and catalog[sku] != metadata:
                    raise ValueError("Inconsistent product metadata across areas.")
                catalog[sku] = metadata
                keys.add(key)
                records.append((sku, product, brand, category, area, int(quantity), cents))
            except (ValueError, InvalidOperation, TypeError) as exc:
                raise ValueError(f"Row {number}: {exc}") from exc
        if not records:
            raise ValueError("Snapshot contains no stock records.")
        # Require an explicit AVAILABLE row so absence cannot be confused with zero.
        missing = sorted(sku for sku in catalog if (sku, "AVAILABLE") not in keys)
        if missing:
            raise ValueError("Each SKU requires an AVAILABLE row, even when zero.")
        return records
    finally:
        workbook.close()


def import_snapshot(conn, source, snapshot_date, *, replace=False):
    """Import an entire dated snapshot atomically; never silently overwrite."""
    snapshot_date = date.fromisoformat(str(snapshot_date)).isoformat()
    records = read_snapshot(source)
    with conn:
        exists = conn.execute("SELECT 1 FROM inventory WHERE snapshot_date=? LIMIT 1", (snapshot_date,)).fetchone()
        if exists and not replace:
            raise ValueError("Snapshot already exists. Explicit replacement required.")
        if exists:
            conn.execute("DELETE FROM inventory WHERE snapshot_date=?", (snapshot_date,))
        conn.executemany("INSERT INTO inventory VALUES (?, ?, ?, ?, ?, ?, ?, ?)", [(snapshot_date, *r) for r in records])
    return len(records)


def dates(conn):
    return [row[0] for row in conn.execute("SELECT DISTINCT snapshot_date FROM inventory ORDER BY snapshot_date")]


def sku_report(conn, day):
    return [dict(row) for row in conn.execute("""
        SELECT sku, product, brand, category,
          SUM(CASE WHEN area='AVAILABLE' THEN quantity ELSE 0 END) AS available,
          SUM(CASE WHEN area='RECEIVING' THEN quantity ELSE 0 END) AS receiving,
          SUM(CASE WHEN area='DAMAGED' THEN quantity ELSE 0 END) AS damaged,
          SUM(quantity) AS total_units,
          SUM(quantity * unit_cost_cents) AS stock_value_cents,
          SUM(CASE WHEN area='AVAILABLE' THEN quantity * unit_cost_cents ELSE 0 END) AS available_value_cents
        FROM inventory WHERE snapshot_date=?
        GROUP BY sku, product, brand, category ORDER BY sku
    """, (str(day),))]


def summary(conn, day):
    rows = sku_report(conn, day)
    return {
        "sku_count": len(rows),
        "available_units": sum(r["available"] for r in rows),
        "receiving_units": sum(r["receiving"] for r in rows),
        "damaged_units": sum(r["damaged"] for r in rows),
        "zero_available_skus": sum(r["available"] == 0 for r in rows),
        "stock_value_cents": sum(r["stock_value_cents"] for r in rows),
        "available_value_cents": sum(r["available_value_cents"] for r in rows),
    }


def history(conn):
    return [{"date": day, **summary(conn, day)} for day in dates(conn)]


def write_table(wb, name, rows):
    ws = wb.create_sheet(name)
    if not rows:
        return
    columns = list(rows[0])
    ws.append(columns)
    for row in rows:
        for index, column in enumerate(columns, 1):
            value = row[column]
            cell = ws.cell(ws.max_row + 1 if index == 1 else ws.max_row, index, value)
            if isinstance(value, str):
                cell.data_type = "s"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1D3557")
    for i, column in enumerate(columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = min(42, max(14, len(column) + 3))


def export_report(conn, day):
    """Return an Excel report, with BRL amounts converted from integer cents."""
    wb = Workbook()
    wb.remove(wb.active)
    def money(row):
        return {k.removesuffix('_cents') + '_brl' if k.endswith('_cents') else k: v / 100 if k.endswith('_cents') else v for k, v in row.items()}
    rows = sku_report(conn, day)
    write_table(wb, "Summary", [money({"date": str(day), **summary(conn, day)})])
    write_table(wb, "Stock by SKU", [money(r) for r in rows])
    write_table(wb, "Zero availability", [money(r) for r in rows if r["available"] == 0])
    write_table(wb, "History", [money(r) for r in history(conn)])
    output = BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue()
