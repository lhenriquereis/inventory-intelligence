"""Generate fully synthetic stock reports; no external data or APIs."""
from datetime import date, timedelta
from pathlib import Path
from random import Random
from contextlib import closing
from openpyxl import Workbook
from src.inventory import ROOT, HEADERS, AREAS, connect, dates, import_snapshot

SAMPLE_DIR = ROOT / "data" / "samples"


def generate_samples(directory=SAMPLE_DIR, days=14, products=30):
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    rng = Random(42)
    catalog = [(f"SKU-{i:04d}", f"Demo Product {i:02d}", f"Demo Brand {1+i%4}", ("Home", "Office", "Personal Care")[i%3], rng.randint(300, 8000)/100) for i in range(1, products+1)]
    paths = []
    for offset in range(days):
        day = date(2026, 11, 1) + timedelta(days=offset)
        wb = Workbook()
        ws = wb.active
        ws.title = "Stock"
        ws.append(HEADERS)
        for i, (sku, product, brand, category, cost) in enumerate(catalog):
            quantities = (0 if (i+offset)%11 == 0 else max(0, 65 + i*3 - offset*3 + rng.randint(-12, 12)), rng.randint(0, 35) if (i+offset)%4 == 0 else 0, rng.randint(0, 4) if i%7 == 0 else 0)
            for area, qty in zip(AREAS, quantities):
                ws.append((sku, product, brand, category, area, qty, cost))
        path = directory / f"stock_{day.isoformat()}.xlsx"
        wb.save(path)
        wb.close()
        paths.append((day.isoformat(), path))
    return paths


def seed_demo(db_path=None):
    with closing(connect() if db_path is None else connect(db_path)) as conn:
        known = set(dates(conn))
        for day, path in generate_samples():
            if day not in known:
                import_snapshot(conn, path, day)
        return len(dates(conn))


if __name__ == "__main__":
    print(f"Demo ready. {seed_demo()} snapshot dates in local SQLite database.")
