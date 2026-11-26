"""Local-only inventory analytics dashboard."""
from contextlib import closing
from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st
from src.inventory import connect, dates, import_snapshot, sku_report, summary, history, export_report
from src.generate_samples import generate_samples, seed_demo

st.set_page_config(page_title="Inventory Intelligence", page_icon="📦", layout="wide")
st.title("Inventory Intelligence")
st.caption("Daily stock snapshots · SQLite analytics · Synthetic demo data · BRL")

with st.sidebar:
    st.header("Data management")
    st.caption("Use fictional data only. This portfolio app stores snapshots locally.")
    if st.button("Load synthetic demo", type="primary"):
        seed_demo()
        st.success("Demo snapshots loaded; existing dates were preserved.")
    # Generate a downloadable template only on request, without changing the database.
    if st.button("Prepare example template"):
        from tempfile import TemporaryDirectory
        with TemporaryDirectory() as temp:
            _, path = generate_samples(temp, days=1, products=3)[0]
            st.session_state["template"] = path.read_bytes()
    if "template" in st.session_state:
        st.download_button("Download example XLSX", st.session_state["template"], "stock_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with st.form("import"):
        upload = st.file_uploader("Upload synthetic snapshot", type=["xlsx"])
        day = st.date_input("Snapshot date", value=date.today())
        replace = st.checkbox("Replace an existing snapshot for this date")
        authorized = st.checkbox("I confirm this file contains only synthetic data")
        submitted = st.form_submit_button("Import snapshot")
    if submitted:
        if upload is None or not authorized:
            st.warning("Select an XLSX and confirm synthetic data.")
        else:
            try:
                with closing(connect()) as conn:
                    count = import_snapshot(conn, BytesIO(upload.getvalue()), day, replace=replace)
                st.success(f"Imported {count} stock-area rows.")
            except Exception as exc:
                st.error(f"Import rejected: {exc}")

with closing(connect()) as conn:
    available_dates = dates(conn)
    if not available_dates:
        st.info("Click Load synthetic demo in the sidebar to begin.")
        st.stop()
    selected = st.selectbox("Snapshot date", available_dates, index=len(available_dates)-1)
    metrics = summary(conn, selected)
    cols = st.columns(4)
    cols[0].metric("Available units", f"{metrics['available_units']:,}")
    cols[1].metric("Receiving units", f"{metrics['receiving_units']:,}")
    cols[2].metric("Zero-availability SKUs", metrics['zero_available_skus'])
    cols[3].metric("Total stock value", f"R$ {metrics['stock_value_cents']/100:,.2f}")
    st.caption(f"{metrics['sku_count']} SKUs · {metrics['damaged_units']} damaged units · Stock value includes all areas.")
    series = pd.DataFrame(history(conn)).set_index("date")
    st.subheader("Daily stock history")
    st.line_chart(series[["available_units", "receiving_units"]])
    rows = pd.DataFrame(sku_report(conn, selected))
    left, right = st.columns(2)
    with left:
        st.subheader("Available units by brand")
        st.bar_chart(rows.groupby("brand")["available"].sum())
    with right:
        st.subheader("Stock value by category (BRL)")
        st.bar_chart(rows.groupby("category")["stock_value_cents"].sum() / 100)
    st.subheader("Inventory detail")
    brands = st.multiselect("Brand filter", sorted(rows["brand"].unique()))
    only_zero = st.checkbox("Show only zero availability")
    filtered = rows.copy()
    if brands:
        filtered = filtered[filtered["brand"].isin(brands)]
    if only_zero:
        filtered = filtered[filtered["available"] == 0]
    filtered["stock_value_brl"] = filtered.pop("stock_value_cents") / 100
    filtered["available_value_brl"] = filtered.pop("available_value_cents") / 100
    st.dataframe(filtered, hide_index=True)
    st.download_button("Download complete snapshot report (XLSX)", export_report(conn, selected), f"inventory_{selected}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("Download contains the complete selected snapshot and history, not just the filtered view. Zero availability is an observation, not a demand forecast.")
