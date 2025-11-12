import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Mini Client Dashboard", layout="wide")
st.title("Mini Client Dashboard")

# -------------------------------------------------
# Helpers
# -------------------------------------------------
def normalize_stock(s: str) -> str:
    if not s:
        return ""
    s = str(s).strip().upper()
    return s if s.endswith(".CA") else f"{s}.CA"

def export_xlsx(df: pd.DataFrame, filename="export.xlsx", sheet_name="Sheet1"):
    """Export a DataFrame as a single-sheet XLSX with simple numeric formatting."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]
        # Basic number format (commas). We don't force percent to avoid scaling surprises.
        for col in ws.columns:
            for cell in col[1:]:
                if isinstance(cell.value, (int, float)):
                    # Quantity columns look nicer as integers when exact
                    if float(cell.value).is_integer():
                        cell.number_format = "#,##0"
                    else:
                        cell.number_format = "#,##0.00"
    buffer.seek(0)
    st.download_button(
        "📥 Download Excel",
        buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# -------------------------------------------------
# Extraction
# -------------------------------------------------
def extract_client_data(file):
    """
    Extracts:
      - Client: B4 (fallback sheet name)
      - Cash: C27
      - Dividends: C32
      - Stocks: from 'Stocks' block (B=Name, C=Qty, E=Price, H=MV, I=Weight)
      - ICs: scan next <=10 rows after stocks end for Stk-300 (Stream MV) & Stk-302 (Momentum MV)
      - AUM: 'Total Assets' row (col C)
      - Total Cash = Cash + Dividends + Stream(MV)
      - Prices table (unique stock -> latest seen price)
    """
    wb = openpyxl.load_workbook(file, data_only=True)
    out = {}
    all_prices = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        client = (ws["B4"].value or sheet_name).title()

        cash = float(ws["C27"].value or 0)
        dividends = float(ws["C32"].value or 0)

        # Locate 'Stocks' start
        start_row = None
        for r in ws.iter_rows(min_row=1, max_row=100):
            if str(r[0].value).strip().lower() == "stocks":
                start_row = r[0].row + 1
                break

        stock_rows = []
        stream_mv = 0.0
        momentum_mv = 0.0
        last_stock_row = None

        if start_row:
            # Read stocks until grey+empty row
            for r in ws.iter_rows(min_row=start_row):
                cell_b = r[1]
                rgb = getattr(getattr(cell_b.fill, "start_color", None), "rgb", None)
                empty_b = (cell_b.value is None) or (str(cell_b.value).strip() == "")
                if rgb == "FFD3D3D3" and empty_b:
                    last_stock_row = r[0].row
                    break

                name = r[1].value
                qty = r[2].value
                price = r[4].value   # Column E
                mv = r[7].value if len(r) > 7 else 0
                wt = r[8].value if len(r) > 8 else 0
                if not name:
                    continue

                name_str = str(name).strip()
                name_up = name_str.upper()

                # Collect prices
                if isinstance(price, (int, float)):
                    all_prices[normalize_stock(name_str)] = float(price)

                # Ignore ICs here (we'll pick them from the post-stocks scan), take only real stocks
                if name_up not in ("STK-300", "STK-302") and isinstance(qty, (int, float)):
                    stock_rows.append({
                        "Company Name": normalize_stock(name_str),
                        "Quantity": float(qty or 0),
                        "Price": float(price or 0),
                        "Market Value": float(mv or 0),
                        "Weight": wt or 0,
                    })

            # Scan up to 10 rows after stocks for ICs
            if not last_stock_row:
                last_stock_row = start_row
            ic_start, ic_end = last_stock_row + 1, last_stock_row + 11
            for r in ws.iter_rows(min_row=ic_start, max_row=ic_end):
                name_cell = r[1].value
                if not name_cell:
                    continue
                nm = str(name_cell).strip().upper()
                mv = r[7].value if len(r) > 7 else 0
                if nm == "STK-300":
                    stream_mv = float(mv or 0)
                elif nm == "STK-302":
                    momentum_mv = float(mv or 0)

        # AUM
        aum = 0.0
        for r in ws.iter_rows(min_row=1, max_row=100):
            if str(r[0].value).strip().lower() == "total assets":
                aum = float(r[2].value or 0)
                break

        total_cash = cash + dividends + stream_mv

        df_stocks = pd.DataFrame(
            stock_rows,
            columns=["Company Name", "Quantity", "Price", "Market Value", "Weight"]
        )

        out[client] = {
            "data": df_stocks,
            "cash": cash,
            "dividends": dividends,
            "stream_mv": stream_mv,
            "momentum_mv": momentum_mv,
            "total_cash": total_cash,
            "aum": aum,
        }

    # Prices table
    prices_df = pd.DataFrame(sorted(all_prices.items()), columns=["Stock", "Price"]) \
                if all_prices else pd.DataFrame(columns=["Stock", "Price"])
    return out, prices_df

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
    for client, info in sorted(data.items()):
        row = {
            "Client": client,
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

def stock_prices_view(prices_df: pd.DataFrame):
    st.subheader("Stock Prices (from Column E)")
    st.dataframe(prices_df, use_container_width=True, hide_index=True)
    export_xlsx(prices_df, filename="stock_prices.xlsx", sheet_name="Prices")

def positions_view(data):
    """
    Vertical block per client:
      Name | <client>
      NAV | <val>
      Total Cash | <val>
      Stocks | Quantity | MV | Weight
      TICKER | qty | mv | weight
      Momentum | <val>
      (blank row)
    """
    st.subheader("Positions View (Vertical Blocks)")

    rows = []
    for client, info in sorted(data.items()):
        rows.append(["Name", client, None, None])
        rows.append(["NAV", info.get("aum", 0), None, None])
        rows.append(["Total Cash", info.get("total_cash", 0), None, None])
        rows.append(["Stocks", "Quantity", "MV", "Weight"])

        df = info["data"]
        if not df.empty:
            for _, r in df.iterrows():
                rows.append([
                    r["Company Name"],
                    r["Quantity"],
                    r["Market Value"],
                    r["Weight"],
                ])

        rows.append(["Momentum", info.get("momentum_mv", 0), None, None])
        rows.append([None, None, None, None])  # spacer

    # ✅ give the sheet real, unique column names
    positions_df = pd.DataFrame(
        rows, columns=["Item", "Value/Qty", "MV", "Weight"]
    )

    st.dataframe(positions_df, use_container_width=True, hide_index=True)
    export_xlsx(positions_df, filename="positions_vertical.xlsx", sheet_name="Positions")

# -------------------------------------------------
# Main
# -------------------------------------------------
uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("Please upload an Excel file to begin.")
else:
    data, prices_df = extract_client_data(uploaded)
    view = st.selectbox(
        "Select View",
        ["Client View", "Total Portfolio View", "Stock Prices View", "Positions View"]
    )
    if view == "Client View":
        client_view(data)
    elif view == "Total Portfolio View":
        total_portfolio_view(data)
    elif view == "Stock Prices View":
        stock_prices_view(prices_df)
    else:
        positions_view(data)

