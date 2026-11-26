from contextlib import closing
from io import BytesIO
from openpyxl import Workbook, load_workbook
import pytest
from src.inventory import HEADERS, connect, import_snapshot, dates, summary, sku_report, export_report
from src.generate_samples import generate_samples


def workbook(rows, headers=HEADERS):
    wb = Workbook()
    ws = wb.active
    ws.title = "Stock"
    ws.append(headers)
    for row in rows:
        ws.append(row)
    stream = BytesIO()
    wb.save(stream)
    wb.close()
    stream.seek(0)
    return stream


def row(area="AVAILABLE", qty=5, cost=12.50, sku="0001"):
    return (sku, "Demo", "Brand A", "Home", area, qty, cost)


@pytest.fixture
def conn():
    with closing(connect(":memory:")) as db:
        yield db


def test_import_and_kpis(conn):
    assert import_snapshot(conn, workbook([row(), row("RECEIVING", 3), row("DAMAGED", 1), row(qty=0, sku="0002")]), "2026-11-01") == 4
    assert summary(conn, "2026-11-01") == {"sku_count": 2, "available_units": 5, "receiving_units": 3, "damaged_units": 1, "zero_available_skus": 1, "stock_value_cents": 11250, "available_value_cents": 6250}
    assert sku_report(conn, "2026-11-01")[0]["sku"] == "0001"


def test_duplicate_date_requires_consent(conn):
    import_snapshot(conn, workbook([row()]), "2026-11-01")
    with pytest.raises(ValueError, match="already exists"):
        import_snapshot(conn, workbook([row(qty=10)]), "2026-11-01")
    assert summary(conn, "2026-11-01")["available_units"] == 5


def test_replace_and_keep_history(conn):
    import_snapshot(conn, workbook([row()]), "2026-11-01")
    import_snapshot(conn, workbook([row(qty=10)]), "2026-11-02")
    import_snapshot(conn, workbook([row(qty=7)]), "2026-11-01", replace=True)
    assert dates(conn) == ["2026-11-01", "2026-11-02"]
    assert summary(conn, "2026-11-01")["available_units"] == 7
    assert summary(conn, "2026-11-02")["available_units"] == 10


@pytest.mark.parametrize("bad_rows", [[row(qty=-1)], [row(qty=1.2)], [row(cost=1.234)], [row("UNKNOWN")], [row(), row()], [row("RECEIVING")], [row(), row("RECEIVING", cost=9)]])
def test_invalid_snapshot_is_atomic(conn, bad_rows):
    import_snapshot(conn, workbook([row()]), "2026-11-01")
    with pytest.raises(ValueError):
        import_snapshot(conn, workbook(bad_rows), "2026-11-01", replace=True)
    assert summary(conn, "2026-11-01")["available_units"] == 5


def test_export(conn):
    import_snapshot(conn, workbook([row()]), "2026-11-01")
    wb = load_workbook(BytesIO(export_report(conn, "2026-11-01")))
    try:
        assert wb.sheetnames == ["Summary", "Stock by SKU", "Zero availability", "History"]
        assert wb["Stock by SKU"]["A2"].value == "0001"
        assert wb["Stock by SKU"]["I2"].value == 62.5
    finally:
        wb.close()


def test_synthetic_samples(conn, tmp_path):
    for day, path in generate_samples(tmp_path, days=3, products=5):
        assert import_snapshot(conn, path, day) == 15
    assert len(dates(conn)) == 3
    assert summary(conn, dates(conn)[-1])["sku_count"] == 5
