import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Mini Client Dashboard", layout="wide")
st.title("Mini Client Dashboard")

# -------------------------------------------------
# Helper Functions
# -------------------------------------------------
def normalize_stock(s: str) -> str:
    if not s:
        return ""
    s = str(s).strip().upper()
    return s if s.endswith(".CA") else f"{s}.CA"

def export_xlsx(df, filename="export.xlsx", sheet_name="Sheet1"):
    """Exports a DataFrame to XLSX in-memory and provides a Streamlit download button."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]
        # apply number formatting
        for col in ws.columns:
            for cell in col[1:]:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"
    buffer.seek(0)
    st.download_button(
        "📥 Download Excel",
        buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# -------------------------------------------------
# Extractor
# -------------------------------------------------
def extract_client_data(file):
    wb = openpyxl.load_workbook(file, data_only=True)
    out = {}
    all_prices = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        client_name = (ws["B4"].value or sheet_name).title()

        cash = float(ws["C27"].value or 0)
        dividends = float(ws["C32"].value or 0)

        # Locate start of "Stocks"
        start_row = None
        for row in ws.iter_rows(min_row=1, max_row=100):
            if str(row[0].value).strip().lower() == "stocks":
                start_row = row[0].row + 1
                break

        stocks, stream_mv, momentum_mv = [], 0.0, 0.0
        last_stock_row = None

        if start_row:
            # Read stocks
            for row in ws.iter_rows(min_row=start_row):
                cell_b = row[1]
                rgb = getattr(getattr(cell_b.fill, "start_color", None), "rgb", None)
                empty_b = (cell_b.value is None) or (str(cell_b.value).strip() == "")
                if rgb == "FFD3D3D3" and empty_b:
                    last_stock_row = row[0].row
                    break

                name = row[1].value
                qty = row[2].value
                price = row[4].value
                mv = row[7].value if len(row) > 7 else 0
                wt = row[8].value if len(row) > 8 else 0
                if not name:
                    continue
                name_str = str(name).strip()
                name_upper = name_str.upper()

                # Record prices globally
                if price is not None and isinstance(price, (int, float)):
                    all_prices[normalize_stock(name_str)] = float(price)

                # Identify ICs
                if name_upper == "STK-300":
                    stream_mv = float(mv or 0)
                elif name_upper == "STK-302":
                    momentum_mv = float(mv or 0)
                else:
                    if isinstance(qty, (int, float)):
                        stocks.append({
                            "Company Name": normalize_stock(name_str),
                            "Quantity": qty,
                            "Price": price or 0,
                            "Market Value": mv or 0,
                            "Weight": wt or 0,
                        })

            # Find ICs 10 rows after stocks
            if not last_stock_row:
                last_stock_row = start_row
            ic_start, ic_end = last_stock_row + 1, last_stock_row + 11
            for row in ws.iter_rows(min_row=ic_start, max_row=ic_end):
                name_cell = row[1].value
                if not name_cell:
                    continue
                name_upper = str(name_cell).strip().upper()
                mv = row[7].value if len(row) > 7 else 0
                if name_upper == "STK-300":
                    stream_mv = float(mv or 0)
                elif name_upper == "STK-302":
                    momentum_mv = float(mv or 0)

        # AUM
        aum = 0.0
        for row in ws.iter_rows(min_row=1, max_row=100):
            if str(row[0].value).strip().lower() == "total assets":
                aum = float(row[2].value or 0)
                break

        total_cash = cash + dividends + stream_mv

        df_stocks = pd.DataFrame(stocks, columns=["Company Name", "Quantity", "Price", "Market Value", "Weight"])

        out[client_name] = {
            "data": df_stocks,
            "cash": cash,
            "dividends": dividends,
            "stream_mv": stream_mv,
            "momentum_mv": momentum_mv,
            "total_cash": total_cash,
            "aum": aum,
        }

    # Compile global prices table
    df_prices = (
        pd.DataFrame(sorted(all_prices.items()), columns=["Stock", "Price"])
        if all_prices else pd.DataFrame(columns=["Stock", "Price"])
    )
    return out, df_prices

# -------------------------------------------------
# Views
# -------------------------------------------------
def client_view(data):
    client = st.selectbox("Select Client", sorted(data.keys()))
    info = data[client]

    st.subheader(client)
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Cash (C27)", f"{info['cash']:,.2f}")
    c2.metric("Dividends (C32)", f"{info['dividends']:,.2f}")
    c3.metric("Stream (Stk-300)", f"{info['stream_mv']:,.2f}")
    c4.metric("Momentum (Stk-302)", f"{info['momentum_mv']:,.2f}")
    c5.metric("Total Cash", f"{info['total_cash']:,.2f}")
    st.metric("AUM", f"{info['aum']:,.2f}")

    st.markdown("**Stock Holdings**")
    st.dataframe(info["data"], use_container_width=True, hide_index=True)
    export_xlsx(info["data"], filename=f"{client}_holdings.xlsx", sheet_name="Holdings")

def total_portfolio_view(data):
    st.subheader("Total Portfolio View")

    all_stocks = sorted({s for v in data.values() for s in v["data"]["Company Name"].unique()})
    rows = []
    for c, info in sorted(data.items()):
        row = {
            "Client": c,
            "Cash": info["cash"],
            "Dividends": info["dividends"],
            "Stream": info["stream_mv"],
            "Momentum": info["momentum_mv"],
            "Total Cash": info["total_cash"],
            "NAV": info["aum"],
        }
        for s in all_stocks:
            row[s] = 0
        for _, r in info["data"].iterrows():
            row[r["Company Name"]] = r["Quantity"]
        rows.append(row)

    cols = ["Client", "Cash", "Dividends", "Stream", "Momentum", "Total Cash"] + all_stocks + ["NAV"]
    mat = pd.DataFrame(rows, columns=cols)
    st.dataframe(mat, use_container_width=True)
    export_xlsx(mat, filename="total_portfolio.xlsx", sheet_name="Portfolio")

def stock_prices_view(df_prices):
    st.subheader("Stock Prices (from Column E)")
    st.dataframe(df_prices, use_container_width=True, hide_index=True)
    export_xlsx(df_prices, filename="stock_prices.xlsx", sheet_name="Prices")

# -------------------------------------------------
# Main App
# -------------------------------------------------
uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("Please upload an Excel file to begin.")
else:
    data, df_prices = extract_client_data(uploaded)
    view = st.selectbox("Select View", ["Client View", "Total Portfolio View", "Stock Prices View"])
    if view == "Client View":
        client_view(data)
    elif view == "Total Portfolio View":
        total_portfolio_view(data)
    else:
        stock_prices_view(df_prices)

