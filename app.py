import streamlit as st
import new_supplier_logic
import pandas as pd
import openpyxl
import os
import zipfile
import io
import re
from docx import Document
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

st.set_page_config(page_title="Auto Docs", layout="wide")
st.title("📦 Phần Mềm Tự Động Hóa Chứng Từ Nhập Khẩu")
st.markdown("---")

st.sidebar.header("Dữ liệu đầu vào")

exchange_rate = 1.0
exchange_rate = st.sidebar.number_input("Tỷ giá RMB sang USD", value=6.74, format="%.4f")
ncc_file = st.sidebar.file_uploader("Kéo thả file Packing List và Invoice (NCC) vào đây (File .xlsx)", type=["xlsx"], accept_multiple_files=True)

def get_prefix(model):
    if "-" in model:
        parts = model.split("-")
        if len(parts[1]) > 0:
            return parts[0] + "-" + parts[1][0]
    import re
    m = re.match(r"([A-Z]+)\d+", model)
    if m: return m.group(1)
    return model

def get_hs_info(inch):
    cm = inch * 2.54
    if cm < 60: return "401039", "UNDER 60CM AND OVER 240CM"
    elif cm <= 180: return "401032", "FROM 60CM TO 180CM"
    elif cm <= 240: return "401034", "FROM 180CM TO 240CM"
    else: return "401039", "UNDER 60CM AND OVER 240CM"

def number_to_words(n):
    units = ["", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE", "TEN",
             "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN"]
    tens = ["", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"]
    n = int(n)
    if n == 0: return "ZERO"
    if n < 20: return units[n]
    if n < 100: return tens[n // 10] + (" " + units[n % 10] if (n % 10 != 0) else "")
    if n < 1000: return units[n // 100] + " HUNDRED" + (" AND " + number_to_words(n % 100) if (n % 100 != 0) else "")
    if n < 1000000: return number_to_words(n // 1000) + " THOUSAND" + (" " + number_to_words(n % 1000) if (n % 1000 != 0) else "")
    return str(n)

def float_to_currency_words(amt):
    amt = round(amt, 2)
    dollars = int(amt)
    cents = int(round((amt - dollars) * 100))
    d_words = number_to_words(dollars)
    c_words = number_to_words(cents)
    res = f"{d_words} DOLLAR{'S' if dollars != 1 else ''}"
    if cents > 0: res += f" AND {c_words} CENT{'S' if cents != 1 else ''}"
    return res + " ONLY."


def normalize_model(m):
    return str(m).upper().replace("-", "").replace(" ", "").strip()

def extract_brand(text):
    t = str(text).upper()
    if "TAKA PRO" in t: return "TAKAPRO"
    if "TAKA" in t: return "TAKA"
    if "YAMATACHI" in t: return "YAMATACHI"
    return None

def get_carton_allocations(df_pl):
    current_ctns = []
    assigned_ctns = set()
    model_cartons = {}
    current_brand = "UNKNOWN"
    
    for idx, row in df_pl.iterrows():
        row_str = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
        b = extract_brand(row_str)
        if b: current_brand = b
            
        col_a = str(row[0]).strip().upper()
        col_b = str(row[1]).strip().upper()
        model = normalize_model(col_b)
        
        if pd.isna(row[1]) or col_b in ["NAN", "NONE", ""]: continue
        if col_b.startswith("RUBBER BELT"): continue
        
        if re.match(r"^\d+$", col_a):
            current_ctns = [int(col_a)]
        elif re.match(r"^\d+-\d+$", col_a):
            parts = col_a.split("-")
            current_ctns = list(range(int(parts[0]), int(parts[1])+1))
        elif current_ctns and col_a in ["NAN", "NONE", ""]:
            pass
        else:
            continue
            
        cartons_for_this_model = 0
        for c in current_ctns:
            if c not in assigned_ctns:
                cartons_for_this_model += 1
                assigned_ctns.add(c)
                
        key = current_brand + "_" + model
        if key not in model_cartons:
            model_cartons[key] = 0
        model_cartons[key] += cartons_for_this_model
        
    return model_cartons, len(assigned_ctns)

def get_pl_groups(df_pl):
    items = []
    current_brand = "UNKNOWN"
    current_ctns = []
    
    for idx, row in df_pl.iterrows():
        col_a = str(row[0]).strip().upper()
        col_b = str(row[1]).strip().upper()
        
        row_str = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
        b = extract_brand(row_str)
        if b: current_brand = b
            
        if pd.isna(row[1]) or col_b in ["NAN", "NONE", ""]: continue
        if col_b.startswith("RUBBER BELT"): continue
        
        if re.match(r"^\d+$", col_a):
            current_ctns = [int(col_a)]
        elif re.match(r"^\d+-\d+$", col_a):
            parts = col_a.split("-")
            current_ctns = list(range(int(parts[0]), int(parts[1])+1))
        elif current_ctns and col_a in ["NAN", "NONE", ""]:
            pass
        else:
            continue
            
        num_c = len(current_ctns)
        inch = row[2] if len(row) > 2 and pd.notna(row[2]) else ""
        qty = float(row[3]) / num_c if len(row) > 3 and pd.notna(row[3]) else 0
        nw = float(row[4]) / num_c if len(row) > 4 and pd.notna(row[4]) else 0
        gw = float(row[5]) / num_c if len(row) > 5 and pd.notna(row[5]) else 0
        
        for c in current_ctns:
            items.append({
                "ctn": c,
                "brand": current_brand,
                "model": col_b,
                "inch": inch,
                "qty": qty,
                "nw": nw,
                "gw": gw
            })

    df_items = pd.DataFrame(items)
    if len(df_items) == 0: return []
    
    carton_counts = df_items.groupby("ctn").size()
    mixed_cartons = set(carton_counts[carton_counts > 1].index)
    
    groups = []
    current_group = None
    
    for idx, row in df_items.iterrows():
        if current_group is None:
            current_group = {
                "start_ctn": row["ctn"],
                "end_ctn": row["ctn"],
                "brand": row["brand"],
                "model": row["model"],
                "inch": row["inch"],
                "total_qty": row["qty"],
                "total_nw": row["nw"],
                "total_gw": row["gw"],
                "cartons_set": {row["ctn"]}
            }
        else:
            num_cartons = len(current_group["cartons_set"])
            curr_qty = round(current_group["total_qty"] / num_cartons, 4) if num_cartons > 0 else 0
            curr_nw = round(current_group["total_nw"] / num_cartons, 4) if num_cartons > 0 else 0
            curr_gw = round(current_group["total_gw"] / num_cartons, 4) if num_cartons > 0 else 0
            
            is_same_specs = (round(row["qty"], 4) == curr_qty and
                             round(row["nw"], 4) == curr_nw and
                             round(row["gw"], 4) == curr_gw)
                             
            is_same_item = (row["brand"] == current_group["brand"] and 
                            row["model"] == current_group["model"] and
                            is_same_specs)
            
            group_it = False
            if is_same_item:
                if row["ctn"] == current_group["end_ctn"]:
                    group_it = True
                elif row["ctn"] == current_group["end_ctn"] + 1:
                    if row["ctn"] not in mixed_cartons and current_group["end_ctn"] not in mixed_cartons:
                        group_it = True
            
            if group_it:
                current_group["end_ctn"] = row["ctn"]
                current_group["total_qty"] += row["qty"]
                current_group["total_nw"] += row["nw"]
                current_group["total_gw"] += row["gw"]
                current_group["cartons_set"].add(row["ctn"])
            else:
                groups.append(current_group)
                current_group = {
                    "start_ctn": row["ctn"],
                    "end_ctn": row["ctn"],
                    "brand": row["brand"],
                    "model": row["model"],
                    "inch": row["inch"],
                    "total_qty": row["qty"],
                    "total_nw": row["nw"],
                    "total_gw": row["gw"],
                    "cartons_set": {row["ctn"]}
                }
    if current_group:
        groups.append(current_group)
        
    return groups

def apply_border_to_merged_cell(ws, start_row, start_col, end_row, end_col):
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            ws.cell(row=r, column=c).border = thin_border

if ncc_file:
    
    st.write("### Chọn Nhà Cung Cấp để xử lý:")
    col1, col2 = st.columns(2)
    with col1:
        btn_feilizhou = st.button("🚀 Xử lý FEILIZHOU", type="primary", use_container_width=True)
    with col2:
        btn_ncc_moi = st.button("🚀 Xử lý Taizhou Norton", type="primary", use_container_width=True)

    supplier = None
    if btn_feilizhou: supplier = "Feilizhou"
    if btn_ncc_moi: supplier = "Taizhou Norton"

    if supplier is not None:

        with st.spinner("Đang tự động tính toán phân bổ thùng và tạo file..."):
            try:
                if isinstance(ncc_file, list):
                    if len(ncc_file) == 1:
                        df_pl_raw = pd.read_excel(ncc_file[0], sheet_name="Packing List", header=None, dtype=str)
                        df_inv_raw = pd.read_excel(ncc_file[0], sheet_name="Invoice", header=None, dtype=str)
                        norton_pl_file = ncc_file[0]
                    else:
                        for f in ncc_file:
                            if "PACKING" in f.name.upper(): 
                                df_pl_raw = pd.read_excel(f, header=None, dtype=str)
                                norton_pl_file = f
                            elif "INVOICE" in f.name.upper(): 
                                df_inv_raw = pd.read_excel(f, header=None, dtype=str)
                else:
                    df_pl_raw = pd.read_excel(ncc_file, sheet_name="Packing List", header=None, dtype=str)
                    df_inv_raw = pd.read_excel(ncc_file, sheet_name="Invoice", header=None, dtype=str)
                    norton_pl_file = ncc_file
                
                model_cartons, total_cartons = get_carton_allocations(df_pl_raw)
                pl_groups = get_pl_groups(df_pl_raw)
                
                inv_no, inv_date = "UNKNOWN", "UNKNOWN"
                for r in range(min(15, len(df_inv_raw))):
                    for c in range(min(10, df_inv_raw.shape[1])):
                        val = str(df_inv_raw.iloc[r, c]).strip().upper()
                        if "INVOICE NO" in val:
                            inv_no = str(df_inv_raw.iloc[r, c]).split(":")[-1].strip()
                            if not inv_no and c+1 < df_inv_raw.shape[1]: inv_no = str(df_inv_raw.iloc[r, c+1]).strip()
                        if "DATE" in val:
                            inv_date = str(df_inv_raw.iloc[r, c]).split(":")[-1].strip()
                            if not inv_date and c+1 < df_inv_raw.shape[1]: inv_date = str(df_inv_raw.iloc[r, c+1]).strip()

                if inv_no.upper() == "INVOICE NO." or inv_no == "": inv_no = "FLZ_INV"
                if inv_date.upper() == "DATE:" or inv_date == "": inv_date = "TODAY"

                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    if supplier == "Norton":
                        import norton_logic
                        norton_logic.process(zip_file, df_pl_raw, df_inv_raw, norton_pl_file, inv_no, inv_date, exchange_rate, extract_brand, normalize_model, get_prefix, get_hs_info, model_cartons, pl_groups, float_to_currency_words, number_to_words)
                    elif supplier == "Taizhou Norton":
                        import new_supplier_logic
                        new_supplier_logic.process(zip_file, df_pl_raw, df_inv_raw, norton_pl_file, inv_no, inv_date, exchange_rate, extract_brand, normalize_model, get_prefix, get_hs_info, None, None, float_to_currency_words, number_to_words)
                    elif supplier == "Feilizhou":
                        
                        ci_path = os.path.join("Templates", "TEMPLATE_CI.xlsx")
                        if os.path.exists(ci_path):
                            wb_ci = openpyxl.load_workbook(ci_path)
                            ws_ci = wb_ci.active
                            ws_ci["J3"] = f"INVOICE NO.: {inv_no}"
                            ws_ci["J4"] = f"DATE: {inv_date}"
                            
                            items_flat = []
                            current_brand = "UNKNOWN"
                            current_ctns = []
                            for idx, row in df_pl_raw.iterrows():
                                row_str = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
                                b = extract_brand(row_str)
                                if b: current_brand = b
                                col_a = str(row[0]).strip().upper()
                                col_b = str(row[1]).strip().upper()
                                if pd.isna(row[1]) or col_b in ["NAN", "NONE", ""]: continue
                                if col_b.startswith("RUBBER BELT"): continue
                                if re.match(r"^\d+$", col_a): current_ctns = [int(col_a)]
                                elif re.match(r"^\d+-\d+$", col_a):
                                    parts = col_a.split("-")
                                    current_ctns = list(range(int(parts[0]), int(parts[1])+1))
                                elif current_ctns and col_a in ["NAN", "NONE", ""]: pass
                                else: continue
                                num_c = len(current_ctns)
                                qty_val = float(row[3]) / num_c if len(row) > 3 and pd.notna(row[3]) else 0
                                for c in current_ctns:
                                    items_flat.append({"ctn": c, "brand": current_brand, "model": col_b, "qty": qty_val})
    
                            df_items_ci = pd.DataFrame(items_flat)
                            assigned_ctns_ci = set()
                            if len(df_items_ci) > 0:
                                for idx, row in df_items_ci.iterrows():
                                    c = row["ctn"]
                                    if c not in assigned_ctns_ci:
                                        df_items_ci.at[idx, "assigned_ctns"] = 1
                                        assigned_ctns_ci.add(c)
                                    else:
                                        df_items_ci.at[idx, "assigned_ctns"] = 0
    
                            df_items_ci["norm_model"] = df_items_ci["model"].apply(normalize_model)
                            grouped = df_items_ci.groupby(["brand", "norm_model"]).agg({"qty": "sum", "assigned_ctns": "sum"}).reset_index()
                            model_data = {}
                            for _, r in grouped.iterrows():
                                key = r["brand"] + "_" + r["norm_model"]
                                model_data[key] = {"qty": r["qty"], "ctns": r["assigned_ctns"]}
    
                            prices = {}
                            for r in range(13, ws_ci.max_row):
                                m = str(ws_ci.cell(row=r, column=3).value).strip()
                                p = str(ws_ci.cell(row=r, column=9).value)
                                if m and m not in ["NAN", "NONE", ""] and not m.startswith("="):
                                    row_str = " ".join([str(ws_ci.cell(row=r, column=c).value).upper() for c in range(1, 10)])
                                    b = extract_brand(row_str)
                                    if b:
                                        try: prices[f"{b}_{normalize_model(m)}"] = float(p)
                                        except: pass
                                        
                            def number_to_words(n):
                                units = ["", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE", "TEN",
                                         "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN"]
                                tens = ["", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"]
                                n = int(n)
                                if n == 0: return "ZERO"
                                if n < 20: return units[n]
                                if n < 100: return tens[n // 10] + (" " + units[n % 10] if (n % 10 != 0) else "")
                                if n < 1000: return units[n // 100] + " HUNDRED" + (" AND " + number_to_words(n % 100) if (n % 100 != 0) else "")
                                if n < 1000000: return number_to_words(n // 1000) + " THOUSAND" + (" " + number_to_words(n % 1000) if (n % 1000 != 0) else "")
                                return str(n)
    
                            def float_to_currency_words(amt):
                                amt = round(amt, 2)
                                dollars = int(amt)
                                cents = int(round((amt - dollars) * 100))
                                d_words = number_to_words(dollars)
                                c_words = number_to_words(cents)
                                res = f"{d_words} DOLLAR{'S' if dollars != 1 else ''}"
                                if cents > 0: res += f" AND {c_words} CENT{'S' if cents != 1 else ''}"
                                return res + " ONLY."
    
                            rows_to_delete_ci = []
                            grand_qty_ci = 0
                            grand_ctns_ci = 0
                            grand_amt_ci = 0
                            total_row_idx = -1
                            
                            for r in range(13, ws_ci.max_row + 1):
                                col_a_val = str(ws_ci.cell(row=r, column=1).value).strip().upper()
                                if col_a_val == "TOTAL":
                                    total_row_idx = r
                                    break
                                
                                if col_a_val.isdigit():
                                    m = str(ws_ci.cell(row=r, column=3).value).strip()
                                    row_str = " ".join([str(ws_ci.cell(row=r, column=c).value).upper() for c in range(1, 10)])
                                    b = extract_brand(row_str)
                                    key = f"{b}_{normalize_model(m)}"
                                    
                                    qty = model_data.get(key, {}).get("qty", 0)
                                    ctns = model_data.get(key, {}).get("ctns", 0)
                                    
                                    if qty == 0:
                                        rows_to_delete_ci.append(r)
                                    else:
                                        price = prices.get(key, 0)
                                        amt = qty * price
                                        ws_ci.cell(row=r, column=6).value = qty
                                        if ctns > 0 and qty % ctns == 0:
                                            ws_ci.cell(row=r, column=7).value = int(qty/ctns)
                                        else:
                                            ws_ci.cell(row=r, column=7).value = ""
                                        ws_ci.cell(row=r, column=8).value = ctns
                                        ws_ci.cell(row=r, column=10).value = amt
                                        
                                        grand_qty_ci += qty
                                        grand_ctns_ci += ctns
                                        grand_amt_ci += amt
    
                            for r in reversed(rows_to_delete_ci):
                                ws_ci.delete_rows(r, 1)
                                if total_row_idx != -1: total_row_idx -= 1
                                
                            if total_row_idx != -1:
                                ws_ci.cell(row=total_row_idx, column=6).value = grand_qty_ci
                                ws_ci.cell(row=total_row_idx, column=8).value = grand_ctns_ci
                                ws_ci.cell(row=total_row_idx, column=10).value = round(grand_amt_ci, 2)
                                
                                for check_r in range(total_row_idx, total_row_idx + 5):
                                    val = str(ws_ci.cell(row=check_r, column=1).value).strip()
                                    if "SAY" in val:
                                        ws_ci.cell(row=check_r, column=1).value = "SAY IN WORD: " + float_to_currency_words(grand_amt_ci)
                                        break
                                        
                            item_idx = 1
                            for r in range(13, ws_ci.max_row + 1):
                                val = str(ws_ci.cell(row=r, column=1).value).strip()
                                if val == "TOTAL": break
                                if val.isdigit() or val != "None":
                                    if ws_ci.cell(row=r, column=3).value:
                                        ws_ci.cell(row=r, column=1).value = item_idx
                                        item_idx += 1
    
                            ci_bytes = io.BytesIO()
                            wb_ci.save(ci_bytes)
                            zip_file.writestr(f"{inv_no}_COMMERCIAL_INVOICE_TAX.xlsx", ci_bytes.getvalue())
                            
                        pl_path = os.path.join("Templates", "TEMPLATE_PL.xlsx")
                        if os.path.exists(pl_path):
                            wb_pl = openpyxl.load_workbook(pl_path)
                            ws_pl = wb_pl.active
                            ws_pl["K3"] = f"INVOICE NO.: {inv_no}"
                            ws_pl["K4"] = f"DATE: {inv_date}"
                            
                            ws_pl.delete_rows(11, ws_pl.max_row - 10)
                            
                            # Fix Ghost Merge Ranges
                            ranges_to_remove = []
                            for rng in ws_pl.merged_cells.ranges:
                                if rng.min_row >= 11:
                                    ranges_to_remove.append(rng)
                            for rng in ranges_to_remove:
                                ws_pl.merged_cells.remove(rng)
                            
                            thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                            center_align = Alignment(horizontal="center", vertical="center")
                            left_align = Alignment(horizontal="left", vertical="center")
                            
                            row_idx = 11
                            current_brand = None
                            pl_assigned_ctns = set()
                            
                            for g in pl_groups:
                                if g["brand"] != current_brand:
                                    current_brand = g["brand"]
                                    display_brand = current_brand
                                    if current_brand == "TAKAPRO": display_brand = "TAKA PRO"
                                    
                                    ws_pl.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=11)
                                    cell = ws_pl.cell(row=row_idx, column=1)
                                    cell.value = f"RUBBER BELT (V-RIBBED) {display_brand}"
                                    cell.font = Font(bold=True)
                                    cell.alignment = left_align
                                    for c in range(1, 12):
                                        ws_pl.cell(row=row_idx, column=c).border = thin_border
                                    row_idx += 1
                                    
                                ctn_str = str(int(g["start_ctn"])) if g["start_ctn"] == g["end_ctn"] else f"{int(g['start_ctn'])}-{int(g['end_ctn'])}"
                                
                                ctns_val = 0
                                for c in g["cartons_set"]:
                                    if c not in pl_assigned_ctns:
                                        ctns_val += 1
                                        pl_assigned_ctns.add(c)
                                        
                                num_cartons = len(g["cartons_set"])
                                pcs_val = g["total_qty"] / num_cartons if num_cartons > 0 else 0
                                pcs_ctn = int(pcs_val) if float(pcs_val).is_integer() else round(pcs_val, 2)
                                    
                                display_brand2 = current_brand
                                if current_brand == "TAKAPRO": display_brand2 = "TAKA PRO"
                                    
                                ws_pl.cell(row=row_idx, column=1).value = ctn_str
                                ws_pl.cell(row=row_idx, column=2).value = f"Rubber Belt (V-Ribbed) {display_brand2} {g['model']}"
                                ws_pl.cell(row=row_idx, column=3).value = g["model"]
                                ws_pl.cell(row=row_idx, column=4).value = g["inch"]
                                ws_pl.cell(row=row_idx, column=5).value = g["total_qty"]
                                ws_pl.cell(row=row_idx, column=6).value = pcs_ctn
                                
                                # ONLY write CTNS if ctns_val > 0, else leave blank
                                if ctns_val > 0: ws_pl.cell(row=row_idx, column=7).value = ctns_val
                                
                                nw_val = round(g["total_nw"] / ctns_val, 2) if ctns_val > 0 else 0
                                gw_val = round(g["total_gw"] / ctns_val, 2) if ctns_val > 0 else 0
                                
                                # ONLY write NW and GW if nw_val > 0, else leave blank
                                if nw_val > 0: ws_pl.cell(row=row_idx, column=8).value = nw_val
                                if gw_val > 0: ws_pl.cell(row=row_idx, column=9).value = gw_val
                                
                                tot_nw = round(g["total_nw"], 2)
                                tot_gw = round(g["total_gw"], 2)
                                if tot_nw > 0: ws_pl.cell(row=row_idx, column=10).value = tot_nw
                                if tot_gw > 0: ws_pl.cell(row=row_idx, column=11).value = tot_gw
                                
                                for c in range(1, 12):
                                    cell = ws_pl.cell(row=row_idx, column=c)
                                    cell.border = thin_border
                                    cell.alignment = center_align
                                    
                                row_idx += 1
                                
                            # Merge logic
                            start_merge_row = 11
                            while start_merge_row <= row_idx - 1:
                                val = ws_pl.cell(row=start_merge_row, column=1).value
                                if val is None or str(val).startswith("RUBBER BELT"):
                                    start_merge_row += 1
                                    continue
                                    
                                end_merge_row = start_merge_row
                                while end_merge_row + 1 <= row_idx - 1:
                                    next_val = ws_pl.cell(row=end_merge_row + 1, column=1).value
                                    if next_val == val:
                                        end_merge_row += 1
                                    else:
                                        break
                                        
                                if end_merge_row > start_merge_row:
                                    for merge_col in [1, 7, 8, 9, 10, 11]:
                                        ws_pl.merge_cells(start_row=start_merge_row, start_column=merge_col, end_row=end_merge_row, end_column=merge_col)
                                        apply_border_to_merged_cell(ws_pl, start_merge_row, merge_col, end_merge_row, merge_col)
                                        ws_pl.cell(row=start_merge_row, column=merge_col).alignment = center_align
                                        
                                start_merge_row = end_merge_row + 1
    
                            # TOTAL ROW
                            ws_pl.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=4)
                            cell = ws_pl.cell(row=row_idx, column=1)
                            cell.value = "TOTAL"
                            cell.font = Font(bold=True, color="FFFFFF")
                            cell.alignment = center_align
                            fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
                            
                            ws_pl.cell(row=row_idx, column=5).value = sum(g["total_qty"] for g in pl_groups)
                            ws_pl.cell(row=row_idx, column=7).value = len(pl_assigned_ctns)
                            ws_pl.cell(row=row_idx, column=10).value = sum(g["total_nw"] for g in pl_groups)
                            ws_pl.cell(row=row_idx, column=11).value = sum(g["total_gw"] for g in pl_groups)
                            
                            for c in range(1, 12):
                                cell = ws_pl.cell(row=row_idx, column=c)
                                cell.border = thin_border
                                cell.fill = fill
                                cell.font = Font(bold=True, color="FFFFFF")
                                
                            pl_bytes = io.BytesIO()
                            wb_pl.save(pl_bytes)
                            zip_file.writestr(f"{inv_no}_PACKING_LIST_TAX.xlsx", pl_bytes.getvalue())
                                
                        fe_path = os.path.join("Templates", "TEMPLATE_FE.xlsx")
                        if os.path.exists(fe_path):
                            wb_fe = openpyxl.load_workbook(fe_path)
                            ws_fe = wb_fe.active
                            
                            for r in range(23, 42):
                                for c in range(1, 10):
                                    cell = ws_fe.cell(row=r, column=c)
                                    if type(cell).__name__ != 'MergedCell':
                                        cell.value = None
    
                            def number_to_words(n):
                                units = ["", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE", "TEN",
                                         "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN"]
                                tens = ["", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"]
                                n = int(n)
                                if n == 0: return "ZERO"
                                if n < 20: return units[n]
                                if n < 100: return tens[n // 10] + (" " + units[n % 10] if (n % 10 != 0) else "")
                                if n < 1000:
                                    return units[n // 100] + " HUNDRED" + (" AND " + number_to_words(n % 100) if (n % 100 != 0) else "")
                                if n < 1000000:
                                    return number_to_words(n // 1000) + " THOUSAND" + (" " + number_to_words(n % 1000) if (n % 1000 != 0) else "")
                                return str(n)
    
                            items_flat = []
                            current_brand = "UNKNOWN"
                            current_ctns = []
                            for idx, row in df_pl_raw.iterrows():
                                row_str = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
                                b = extract_brand(row_str)
                                if b: current_brand = b
                                    
                                col_a = str(row[0]).strip().upper()
                                col_b = str(row[1]).strip().upper()
                                
                                if pd.isna(row[1]) or col_b in ["NAN", "NONE", ""]: continue
                                if col_b.startswith("RUBBER BELT"): continue
                                
                                if re.match(r"^\d+$", col_a): current_ctns = [int(col_a)]
                                elif re.match(r"^\d+-\d+$", col_a):
                                    parts = col_a.split("-")
                                    current_ctns = list(range(int(parts[0]), int(parts[1])+1))
                                elif current_ctns and col_a in ["NAN", "NONE", ""]: pass
                                else: continue
                                
                                num_c = len(current_ctns)
                                inch_val = float(row[2]) if len(row) > 2 and pd.notna(row[2]) else 0
                                qty_val = float(row[3]) / num_c if len(row) > 3 and pd.notna(row[3]) else 0
                                
                                for c in current_ctns:
                                    items_flat.append({
                                        "ctn": c,
                                        "brand": current_brand,
                                        "model": col_b,
                                        "inch": inch_val,
                                        "qty": qty_val
                                    })
    
                            df_items = pd.DataFrame(items_flat)
                            assigned_ctns = set()
                            if len(df_items) > 0:
                                for idx, row in df_items.iterrows():
                                    c = row["ctn"]
                                    if c not in assigned_ctns:
                                        df_items.at[idx, "assigned_ctns"] = 1
                                        assigned_ctns.add(c)
                                    else:
                                        df_items.at[idx, "assigned_ctns"] = 0
    
                                def get_prefix(model):
                                    if "-" in model:
                                        parts = model.split("-")
                                        if len(parts[1]) > 0:
                                            return parts[0] + "-" + parts[1][0]
                                    return model
    
                                df_items["prefix"] = df_items["model"].apply(get_prefix)
    
                                def get_hs_info(inch):
                                    cm = inch * 2.54
                                    if cm <= 180: return "401032", "FROM 60CM TO 180CM"
                                    elif cm <= 240: return "401034", "FROM 180CM TO 240CM"
                                    else: return "401039", "OVER 240CM"
    
                                df_items[["hs_code", "size_desc"]] = df_items.apply(lambda r: pd.Series(get_hs_info(r["inch"])), axis=1)
    
                                brand_order = {"TAKAPRO": 1, "TAKA": 2, "YAMATACHI": 3}
                                df_items['brand_order'] = df_items['brand'].map(brand_order).fillna(99)
                                
                                fe_groups = df_items.groupby(["brand_order", "brand", "hs_code", "size_desc"])
    
                                row_idx = 23
                                item_num = 1
                                total_fe_qty = 0
                                
                                for name, group in fe_groups:
                                    _, brand, hs_code, size_desc = name
                                    prefixes = ", ".join(sorted(group["prefix"].unique()))
                                    total_cartons = int(group["assigned_ctns"].sum())
                                    total_qty = int(group["qty"].sum())
                                    total_fe_qty += total_qty
                                    
                                    display_brand = "TAKA PRO" if brand == "TAKAPRO" else brand
                                    
                                    ctn_words = number_to_words(total_cartons)
                                    desc = f"{ctn_words} ({total_cartons}) CTNS OF RUBBER BELT (V-RIBBED) {display_brand} {prefixes} {size_desc}"
                                    
                                    cell1 = ws_fe.cell(row=row_idx, column=1)
                                    if type(cell1).__name__ != 'MergedCell': cell1.value = item_num
                                    
                                    if item_num == 1:
                                        cell2 = ws_fe.cell(row=row_idx, column=2)
                                        if type(cell2).__name__ != 'MergedCell': cell2.value = "TAKA\nYAMATACHI"
                                        cell8 = ws_fe.cell(row=row_idx, column=8)
                                        if type(cell8).__name__ != 'MergedCell': cell8.value = f"{inv_no}\n{inv_date}"
                                        
                                    cell3 = ws_fe.cell(row=row_idx, column=3)
                                    if type(cell3).__name__ != 'MergedCell': cell3.value = desc
                                    
                                    cell5 = ws_fe.cell(row=row_idx, column=5)
                                    if type(cell5).__name__ != 'MergedCell': cell5.value = '\n     “PE”'
                                    
                                    cell6 = ws_fe.cell(row=row_idx, column=6)
                                    if type(cell6).__name__ != 'MergedCell': cell6.value = total_qty
                                    
                                    row_idx += 1
                                    
                                    if item_num == 1:
                                        cell2_2 = ws_fe.cell(row=row_idx, column=2)
                                        if type(cell2_2).__name__ != 'MergedCell': cell2_2.value = "C/NO"
                                        
                                    cell3_2 = ws_fe.cell(row=row_idx, column=3)
                                    if type(cell3_2).__name__ != 'MergedCell': cell3_2.value = f"HS CODE: {hs_code}"
                                    
                                    if row_idx > 40:
                                        for col in range(1, 10):
                                            c_source = ws_fe.cell(row=24, column=col)
                                            c_target = ws_fe.cell(row=row_idx, column=col)
                                            if type(c_target).__name__ != 'MergedCell' and type(c_source).__name__ != 'MergedCell':
                                                if c_source.has_style:
                                                    c_target.font = openpyxl.styles.Font(name=c_source.font.name, size=c_source.font.size, bold=c_source.font.bold)
                                                    c_target.border = openpyxl.styles.Border(left=c_source.border.left, right=c_source.border.right, top=c_source.border.top, bottom=c_source.border.bottom)
                                                    c_target.fill = openpyxl.styles.PatternFill(fill_type=c_source.fill.fill_type, fgColor=c_source.fill.fgColor, bgColor=c_source.fill.bgColor)
                                                    c_target.alignment = openpyxl.styles.Alignment(horizontal=c_source.alignment.horizontal, vertical=c_source.alignment.vertical, wrap_text=c_source.alignment.wrap_text)
                                    
                                    row_idx += 1
                                    item_num += 1
    
                                cell_tot = ws_fe.cell(row=row_idx, column=3)
                                if type(cell_tot).__name__ != 'MergedCell': cell_tot.value = f"TOTAL QUANTITIES: {total_fe_qty:,.0f} PCS"
                                cell_qty = ws_fe.cell(row=row_idx, column=6)
                                if type(cell_qty).__name__ != 'MergedCell': cell_qty.value = total_fe_qty
                                cell_pe = ws_fe.cell(row=row_idx, column=5)
                                if type(cell_pe).__name__ != 'MergedCell': cell_pe.value = '\n     “PE”'
                                
                            fe_bytes = io.BytesIO()
                            wb_fe.save(fe_bytes)
                            zip_file.writestr(f"DRAFT_FE_{inv_no}_TAX.xlsx", fe_bytes.getvalue())
                                
                        contract_path = os.path.join("Templates", "TEMPLATE_CONTRACT.docx")
                        if os.path.exists(contract_path):
                            doc = Document(contract_path)
                            for p in doc.paragraphs:
                                if p.text.startswith("No:"): p.text = f"No: {inv_no}"
                                elif p.text.startswith("Date:"): p.text = f"Date: {inv_date}"
                                
                            table = doc.tables[0]
                            rows_to_delete = []
                            summary_row = None
                            accum_qty = 0
                            accum_ctns = 0
                            
                            grand_qty = 0
                            grand_ctns = 0
                            grand_amt = 0
                            
                            for i, row in enumerate(table.rows[1:]):
                                col0 = row.cells[0].text.strip()
                                if col0 == "TOTAL":
                                    if summary_row is not None:
                                        if accum_qty == 0: rows_to_delete.append(summary_row)
                                        else:
                                            summary_row.cells[4].text = f" {accum_qty:,.0f} "
                                            summary_row.cells[6].text = f" {accum_ctns:,.0f} "
                                    row.cells[4].text = f" {grand_qty:,.0f} "
                                    row.cells[6].text = f" {grand_ctns:,.0f} "
                                    row.cells[8].text = f" $ {grand_amt:,.2f} "
                                    
                                elif "SAY IN WORD:" in col0:
                                    for cell in row.cells:
                                        cell.text = "SAY IN WORD: " + float_to_currency_words(grand_amt)
                                    
                                elif "from " in col0.lower() or "under " in col0.lower() or "over " in col0.lower():
                                    if summary_row is not None:
                                        if accum_qty == 0: rows_to_delete.append(summary_row)
                                        else:
                                            summary_row.cells[4].text = f" {accum_qty:,.0f} "
                                            summary_row.cells[6].text = f" {accum_ctns:,.0f} "
                                    summary_row = row
                                    accum_qty = 0
                                    accum_ctns = 0
                                    
                                elif col0.isdigit():
                                    desc = row.cells[1].text
                                    b = extract_brand(desc)
                                    m = normalize_model(row.cells[3].text)
                                    key = f"{b}_{m}"
                                    
                                    qty = model_data.get(key, {}).get("qty", 0)
                                    ctns = model_data.get(key, {}).get("ctns", 0)
                                    
                                    if qty == 0:
                                        rows_to_delete.append(row)
                                    else:
                                        price = prices.get(key)
                                        if price is None:
                                            price_str = row.cells[7].text
                                            price = float(re.sub(r'[^\d.]', '', price_str))
                                        amt = qty * price
                                        
                                        row.cells[4].text = f" {qty:,.0f} "
                                        if ctns > 0 and qty % ctns == 0:
                                            row.cells[5].text = f" {int(qty/ctns)} "
                                        else:
                                            row.cells[5].text = ""
                                        row.cells[6].text = f" {ctns:,.0f} "
                                        row.cells[8].text = f" $ {amt:,.2f} "
                                        
                                        accum_qty += qty
                                        accum_ctns += ctns
                                        grand_qty += qty
                                        grand_ctns += ctns
                                        grand_amt += amt
    
                            for r in rows_to_delete:
                                table._element.remove(r._element)
    
                            item_idx = 1
                            for i, row in enumerate(table.rows[1:]):
                                col0 = row.cells[0].text.strip()
                                if col0.isdigit():
                                    row.cells[0].text = str(item_idx)
                                    item_idx += 1
                                    
                            contract_bytes = io.BytesIO()
                            doc.save(contract_bytes)
                            zip_file.writestr(f"CONTRACT_{inv_no}.docx", contract_bytes.getvalue())
                            
                st.success(f"✅ Đã xử lý phân bổ thùng thành công! Tổng: {total_cartons} CTNS.")
                st.download_button(
                    label="📥 TẢI VỀ 4 FILE (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"CHUNGTU_{inv_no}.zip",
                    mime="application/zip",
                    type="primary"
                )
            except Exception as e:
                import traceback
                st.error(f"Lỗi: {e}")
                st.code(traceback.format_exc())
