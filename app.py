import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill as XLFill
from openpyxl.utils import get_column_letter
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

def total_portfolio_view_weights(data):
    st.subheader("Total Portfolio View (Weights)")
    all_stocks = sorted({s for v in data.values() for s in v["data"]["Company Name"].unique()})
    rows = []
    for client, info in sorted(data.items()):
        row = {
            "Client": client,
            "Cash": info["cash"],
            "NAV": info["aum"],
        }
        for s in all_stocks:
            row[s] = 0
        for _, r in info["data"].iterrows():
            row[r["Company Name"]] = r["Weight"]
        rows.append(row)
    cols = ["Client", "Cash"] + all_stocks + ["NAV"]
    mat = pd.DataFrame(rows, columns=cols)
    st.dataframe(mat, use_container_width=True)
    export_xlsx(mat, filename="total_portfolio_weights.xlsx", sheet_name="Portfolio")

def stock_prices_view(prices_df: pd.DataFrame):
    st.subheader("Stock Prices (from Column E)")
    st.dataframe(prices_df, use_container_width=True, hide_index=True)
    export_xlsx(prices_df, filename="stock_prices.xlsx", sheet_name="Prices")


from io import BytesIO
import re
import pandas as pd

def positions_view(data, prices_df=None):
    """
    Positions view:
      - Uses a Groups file (Sequence, Groups, Name) to cluster & order clients.
      - Creates ONE sheet per group.
      - Layout per sheet (columns A–D):
          Item | Value/Qty | Weight | MV
          Group, Name, NAV, Total Cash, Stocks header, stock rows, Momentum, blank
      - For each stock row:
          MV     (col D) = Qty * Price (via VLOOKUP from group summary)
          Weight (col C) = MV / NAV
      - Group summary in columns H:...:
          H1: "Group Summary"
          Totals (Cash, NAV) in H2:I3
          H5:L..: Stock, Sum Qty, Sum MV, Weight, Price
      - Streamlit preview shows the combined rows exactly as written (all groups stacked).
    """
    st.subheader("Positions View (Per-Group Sheets + Summary + Formulas)")

    # ---------- 0) Groups mapping upload ----------
    grp_file = st.file_uploader(
        "Upload Groups mapping (columns: Sequence, Groups, Name) — optional",
        type=["xlsx", "csv"],
        key="groups_mapping_upload_positions"
    )

    def _norm_name(x):
        return str(x).strip().title()

    # Base list of clients
    clients_df = pd.DataFrame({"Name": sorted(data.keys())})
    grouped_meta = None

    if grp_file is not None:
        try:
            gdf = (
                pd.read_csv(grp_file)
                if grp_file.name.lower().endswith(".csv")
                else pd.read_excel(grp_file)
            )
            colmap = {c.lower().strip(): c for c in gdf.columns}
            name_col = colmap.get("name")
            seq_col = colmap.get("sequence")
            grp_col = colmap.get("groups") or colmap.get("group")

            if not name_col:
                st.error("Groups file must contain a 'Name' column.")
                return

            gdf = gdf.rename(
                columns={
                    name_col: "Name",
                    (seq_col if seq_col else "Sequence"): "Sequence",
                    (grp_col if grp_col else "Groups"): "Groups",
                }
            )
            gdf["Name"] = gdf["Name"].map(_norm_name)
            if "Groups" not in gdf.columns:
                gdf["Groups"] = "Ungrouped"
            if "Sequence" not in gdf.columns:
                gdf["Sequence"] = None
            gdf["Sequence"] = pd.to_numeric(gdf["Sequence"], errors="coerce")

            clients_df["Name"] = clients_df["Name"].map(_norm_name)
            merged = clients_df.merge(
                gdf[["Name", "Groups", "Sequence"]], on="Name", how="left"
            )
            merged["Groups"] = merged["Groups"].fillna("Ungrouped")
            merged["SeqSort"] = merged["Sequence"].fillna(10**9)
            grouped_meta = (
                merged.sort_values(["Groups", "SeqSort", "Name"])
                .reset_index(drop=True)
            )
        except Exception as e:
            st.warning(f"Could not read groups file: {e}")

        if grouped_meta is None:
            grouped_meta = clients_df.copy()
            grouped_meta["Groups"] = "Ungrouped"
            grouped_meta["SeqSort"] = 0

    if "Sequence" not in grouped_meta.columns:
        grouped_meta["Sequence"] = None

    st.write("**Planned group order:**")
    preview_cols = [c for c in ["Groups","Name","Sequence"] if c in grouped_meta.columns]
    st.dataframe(
        grouped_meta[preview_cols],
        hide_index=True,
        use_container_width=True,
    )

    def sanitize_sheet_name(s: str) -> str:
        s = re.sub(r'[:\\/?*\[\]]', "-", str(s)).strip()
        return (s or "Group")[:31]

    # Pre-build a price map from prices_df if provided
    price_map = {}
    if prices_df is not None and not prices_df.empty:
        if "Stock" in prices_df.columns and "Price" in prices_df.columns:
            price_map = dict(zip(prices_df["Stock"], prices_df["Price"]))

    # Fills
    blue_fill = XLFill(start_color="A7C6ED", end_color="A7C6ED", fill_type="solid")
    gblue_fill = XLFill(start_color="8FB7EA", end_color="8FB7EA", fill_type="solid")
    light_grey = XLFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    dark_grey = XLFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")

    # ---------- Build preview & workbook ----------
    all_preview_rows = []  # for Streamlit preview (all groups stacked)
    preview_header = ["Item", "Value/Qty", "Weight", "MV"]

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        created_any_sheet = False

        for grp_name, sub in grouped_meta.groupby("Groups", sort=False):
            # Filter clients that exist in data
            sub_valid = sub[sub["Name"].isin(data.keys())].copy()

            if sub_valid.empty:
                # create minimal placeholder sheet
                sheet_name = sanitize_sheet_name(grp_name)
                base, idx = sheet_name, 1
                while sheet_name in writer.book.sheetnames:
                    sheet_name = sanitize_sheet_name(f"{base}-{idx}")
                    idx += 1
                pd.DataFrame(
                    [["Group", grp_name, None, None]],
                    columns=preview_header,
                ).to_excel(writer, index=False, sheet_name=sheet_name)
                created_any_sheet = True
                continue

            # ----- Build rows + block markers (per client) -----
            rows = []
            block_markers = []  # list of dict(nav_row, first_stock_row, last_stock_row)
            current_row = 2  # Excel data row (1 is header)

            for _, r in sub_valid.iterrows():
                client = r["Name"]
                info = data[client]

                # Group label row
                rows.append(["Group", grp_name, None, None])
                current_row += 1  # this row will be at Excel row current_row-1? Wait: we add header later; easier: we will recalc below
                # Actually simpler: we don't use this row for formulas so no marker needed.

                # Name row
                name_excel_row = current_row
                rows.append(["Name", client, None, None])
                current_row += 1

                # NAV row
                nav_excel_row = current_row
                nav_val = float(info.get("aum", 0) or 0)
                rows.append(["NAV", nav_val, None, None])
                current_row += 1

                # Total Cash row
                total_cash = float(info.get("total_cash", 0) or 0)
                rows.append(["Total Cash", total_cash, None, None])
                current_row += 1

                # Stocks header row
                rows.append(["Stocks", "Quantity", "Weight", "MV"])
                current_row += 1

                # stock lines
                df = info["data"]
                first_stock_row = current_row
                stock_count = 0
                if not df.empty:
                    for _, srow in df.iterrows():
                        stock = srow.get("Company Name")
                        qty = float(srow.get("Quantity", 0) or 0)
                        # leave Weight & MV to formulas
                        rows.append([stock, qty, None, None])
                        stock_count += 1
                        current_row += 1
                last_stock_row = (
                    first_stock_row + stock_count - 1 if stock_count > 0 else None
                )

                # Momentum row
                rows.append(["Momentum", float(info.get("momentum_mv", 0) or 0), None, None])
                current_row += 1

                # spacer
                rows.append(["", "", "", ""])
                current_row += 1

                block_markers.append(
                    {
                        "nav_row": nav_excel_row,
                        "first_stock_row": first_stock_row,
                        "last_stock_row": last_stock_row,
                    }
                )

            # For preview: accumulate these rows with group separation
            all_preview_rows.extend(rows + [["", "", "", ""]])

            # Write this group sheet
            sheet_name = sanitize_sheet_name(grp_name)
            base, idx = sheet_name, 1
            while sheet_name in writer.book.sheetnames:
                sheet_name = sanitize_sheet_name(f"{base}-{idx}")
                idx += 1

            grp_df = pd.DataFrame(rows, columns=preview_header)
            grp_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0, startcol=0)
            ws = writer.sheets[sheet_name]
            created_any_sheet = True

            # ---------- Group summary (H:...) ----------
            group_clients = sub_valid["Name"].tolist()
            total_nav = sum(float(data[n].get("aum", 0) or 0) for n in group_clients)
            total_cash_sum = sum(float(data[n].get("total_cash", 0) or 0) for n in group_clients)

            stocks_rows = []
            for n in group_clients:
                dfc = data[n]["data"]
                if not dfc.empty:
                    for _, srow in dfc.iterrows():
                        stocks_rows.append(
                            {
                                "Stock": srow.get("Company Name"),
                                "Qty": float(srow.get("Quantity", 0) or 0),
                                "Price": float(srow.get("Price", 0) or 0),
                                "MV": float(srow.get("Market Value", 0) or 0),
                            }
                        )

            sum_df = pd.DataFrame(stocks_rows)
            if not sum_df.empty:
                # prefer MV from Qty * Price if price is nonzero
                sum_df["MV_calc"] = sum_df.apply(
                    lambda row: (row["Qty"] * row["Price"])
                    if row["Price"]
                    else row["MV"],
                    axis=1,
                )
                gsum = sum_df.groupby("Stock", as_index=False).agg(
                    **{"Sum Qty": ("Qty", "sum"), "Sum MV": ("MV_calc", "sum")}
                )
                # group weight
                gsum["Weight"] = gsum["Sum MV"].div(total_nav).fillna(0.0)
                # attach Price: from price_map if available, else implied Sum MV / Sum Qty
                if price_map:
                    gsum["Price"] = gsum["Stock"].map(price_map)
                else:
                    gsum["Price"] = gsum.apply(
                        lambda row: (row["Sum MV"] / row["Sum Qty"])
                        if row["Sum Qty"]
                        else None,
                        axis=1,
                    )
            else:
                gsum = pd.DataFrame(columns=["Stock", "Sum Qty", "Sum MV", "Weight", "Price"])

            # write summary
            ws.cell(row=1, column=8, value="Group Summary")  # H1
            ws.cell(row=2, column=8, value="Total Cash")
            ws.cell(row=2, column=9, value=total_cash_sum)
            ws.cell(row=3, column=8, value="Total NAV")
            ws.cell(row=3, column=9, value=total_nav)
            ws["I2"].number_format = "#,##0.00"
            ws["I3"].number_format = "#,##0.00"

            start_r = 5
            headers = ["Stock", "Sum Qty", "Sum MV", "Weight", "Price"]
            for j, h in enumerate(headers, start=8):  # H..L
                ws.cell(row=start_r, column=j, value=h)

            for i, row in gsum.iterrows():
                rr = start_r + 1 + i
                ws.cell(row=rr, column=8, value=row["Stock"])
                ws.cell(row=rr, column=9, value=row["Sum Qty"])
                ws.cell(row=rr, column=10, value=row["Sum MV"])
                ws.cell(row=rr, column=11, value=row["Weight"])
                ws.cell(row=rr, column=12, value=row["Price"])

                ws[f"I{rr}"].number_format = "#,##0"      # Sum Qty
                ws[f"J{rr}"].number_format = "#,##0.00"   # Sum MV
                ws[f"K{rr}"].number_format = "0.00%"      # Weight
                ws[f"L{rr}"].number_format = "#,##0.00"   # Price

            # price lookup range for formulas (stock vs price)
            if not gsum.empty:
                price_start = start_r + 1
                price_end = start_r + len(gsum)
                # H..L, but for VLOOKUP we really just need Stock (H) + Price (L) in same range
                price_range = f"$H${price_start}:$L${price_end}"
            else:
                price_range = None

            # ---------- Number formats for main table ----------
            for cidx in range(2, 5):  # B=Value/Qty, C=Weight, D=MV
                for r in range(2, ws.max_row + 1):
                    cell = ws.cell(row=r, column=cidx)
                    if isinstance(cell.value, (int, float)):
                        if cidx == 3:  # Weight
                            cell.number_format = "0.00%"
                        else:
                            if float(cell.value).is_integer():
                                cell.number_format = "#,##0"
                            else:
                                cell.number_format = "#,##0.00"

            # ---------- Apply formulas for MV & Weight ----------
            if price_range:
                for bm in block_markers:
                    nav_row = bm["nav_row"]
                    f_s = bm["first_stock_row"]
                    l_s = bm["last_stock_row"]
                    if not l_s:
                        continue
                    nav_val_cell = ws.cell(row=nav_row, column=2)  # NAV value in col B

                    for r in range(f_s, l_s + 1):
                        stock_cell = ws.cell(row=r, column=1)   # A
                        qty_cell = ws.cell(row=r, column=2)     # B
                        wt_cell = ws.cell(row=r, column=3)      # C
                        mv_cell = ws.cell(row=r, column=4)      # D

                        mv_cell.value = (
                            f'=IFERROR({qty_cell.coordinate}*'
                            f'VLOOKUP({stock_cell.coordinate},{price_range},5,FALSE),0)'
                        )
                        mv_cell.number_format = "#,##0.00"

                        wt_cell.value = (
                            f'=IFERROR({mv_cell.coordinate}/{nav_val_cell.coordinate},0)'
                        )
                        wt_cell.number_format = "0.00%"

            # ---------- Styling: group/name/alternating rows ----------
            toggle = True
            for r in range(2, ws.max_row + 1):
                label = (ws.cell(row=r, column=1).value or "").strip().lower()
                if label == "group":
                    for c in range(1, 5):
                        ws.cell(row=r, column=c).fill = gblue_fill
                elif label == "name":
                    for c in range(1, 5):
                        ws.cell(row=r, column=c).fill = blue_fill
                    toggle = True
                elif label == "":
                    continue
                else:
                    fill = light_grey if toggle else dark_grey
                    for c in range(1, 5):
                        ws.cell(row=r, column=c).fill = fill
                    toggle = not toggle

        # if for some reason no sheet got written
        if not created_any_sheet:
            pd.DataFrame(
                [["No data", "No matching clients in groups file"]],
                columns=["Item", "Value"],
            ).to_excel(writer, index=False, sheet_name="Summary")

    buf.seek(0)
    st.download_button(
        "📥 Download Excel (Positions by Group)",
        buf,
        file_name="positions_by_group.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ---------- Streamlit preview mimicking layout ----------
    if all_preview_rows:
        preview_df = pd.DataFrame(all_preview_rows, columns=preview_header)
        st.caption("Preview (structure matches Excel sheets):")
        st.dataframe(preview_df, hide_index=True, use_container_width=True)
    else:
        st.info("No rows to preview (no matching clients).")






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
        ["Client View", "Total Portfolio View", "Total Portfolio View (Weights)", "Stock Prices View", "Positions View"]
    )
    if view == "Client View":
        client_view(data)
    elif view == "Total Portfolio View":
        total_portfolio_view(data)
    elif view == "Stock Prices View":
        stock_prices_view(prices_df)
    elif view == "Total Portfolio View (Weights)":
        total_portfolio_view_weights(data)
    else:
        positions_view(data,prices_df)


