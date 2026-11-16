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

# ----------------- PAGE CONFIG -----------------
st.set_page_config(
    page_title="Tự Động Hóa Chứng Từ Nhập Khẩu",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

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



def load_excel_files(files):
    """
    Tự động nhận diện và đọc file Excel:
    - 1 file có 2 sheet (Packing List & Invoice)
    - hoặc 2 file riêng biệt
    """
    if not files:
        return None, None, None, None, None, None
        
    if len(files) == 1:
        f = files[0]
        try:
            b_data = f.getvalue() if hasattr(f, 'getvalue') else (f.read() if hasattr(f, 'read') else open(f, 'rb').read())
            xl = pd.ExcelFile(io.BytesIO(b_data))
            sheet_names = xl.sheet_names
            pl_sheet = None
            inv_sheet = None
            
            for s in sheet_names:
                s_lower = s.lower().strip()
                if any(k in s_lower for k in ["pack", "pl", "list"]):
                    pl_sheet = s
                elif any(k in s_lower for k in ["inv", "ci", "bill", "commer"]):
                    inv_sheet = s
                    
            if (not pl_sheet or not inv_sheet) and len(sheet_names) == 2:
                s0_text = str(pd.read_excel(io.BytesIO(b_data), sheet_name=sheet_names[0], nrows=15, header=None)).lower()
                if any(k in s0_text for k in ["pack", "ctn", "n.w", "g.w"]):
                    pl_sheet, inv_sheet = sheet_names[0], sheet_names[1]
                else:
                    pl_sheet, inv_sheet = sheet_names[1], sheet_names[0]
                    
            if not pl_sheet or not inv_sheet:
                return None, None, None, None, None, f"File {getattr(f, 'name', 'Excel')} có các sheet: {sheet_names}. Cần có sheet Packing List và Invoice."
                
            df_pl = pd.read_excel(io.BytesIO(b_data), sheet_name=pl_sheet, header=None, dtype=str)
            df_inv = pd.read_excel(io.BytesIO(b_data), sheet_name=inv_sheet, header=None, dtype=str)
            name = getattr(f, 'name', 'File.xlsx')
            return df_pl, df_inv, io.BytesIO(b_data), f"{name} (Sheet: {pl_sheet})", f"{name} (Sheet: {inv_sheet})", None
        except Exception as e:
            return None, None, None, None, None, f"Lỗi đọc file: {e}"
            
    else:
        pl_file = None
        inv_file = None
        for f in files:
            name_upper = getattr(f, 'name', '').upper()
            if any(k in name_upper for k in ["PACK", "PL"]):
                pl_file = f
            elif any(k in name_upper for k in ["INV", "CI", "COMMERCIAL"]):
                inv_file = f
                
        if not pl_file or not inv_file:
            pl_file = files[0]
            inv_file = files[1]
            
        try:
            b_pl = pl_file.getvalue() if hasattr(pl_file, 'getvalue') else (pl_file.read() if hasattr(pl_file, 'read') else open(pl_file, 'rb').read())
            b_inv = inv_file.getvalue() if hasattr(inv_file, 'getvalue') else (inv_file.read() if hasattr(inv_file, 'read') else open(inv_file, 'rb').read())
            df_pl = pd.read_excel(io.BytesIO(b_pl), header=None, dtype=str)
            df_inv = pd.read_excel(io.BytesIO(b_inv), header=None, dtype=str)
            name_pl = getattr(pl_file, 'name', 'Packing_List.xlsx')
            name_inv = getattr(inv_file, 'name', 'Invoice.xlsx')
            return df_pl, df_inv, io.BytesIO(b_pl), name_pl, name_inv, None
        except Exception as e:
            return None, None, None, None, None, f"Lỗi đọc file: {e}"

def process_feilizhou(zip_file, df_pl_raw, df_inv_raw, inv_no, inv_date, exchange_rate, model_cartons=None, pl_groups=None):
    # 1. Parse PL rows
    raw_items = []
    last_ctn = ""
    cur_brd = "TAKA PRO"
    
    for idx, row in df_pl_raw.iterrows():
        c = row[0] if len(row) > 0 and pd.notna(row[0]) else None
        m = row[1] if len(row) > 1 and pd.notna(row[1]) else None
        inch = row[2] if len(row) > 2 and pd.notna(row[2]) else None
        q = row[3] if len(row) > 3 and pd.notna(row[3]) else None
        nw = row[4] if len(row) > 4 and pd.notna(row[4]) else None
        gw = row[5] if len(row) > 5 and pd.notna(row[5]) else None
        
        row_vals = [row[col] for col in range(min(7, len(row))) if pd.notna(row[col])]
        row_str = " ".join([str(x).upper() for x in row_vals])
        
        # Check brand header
        b = extract_brand(row_str)
        if b:
            cur_brd = "TAKA PRO" if b == "TAKAPRO" else b
            
        if not m or str(m).strip().upper() in ["TOTAL", "ITEM"]:
            continue
            
        sm = str(m).strip()
        if (not inch or str(inch).strip() == "") and (not q or str(q).strip() == "") and "-" not in sm:
            continue
            
        try:
            qty_num = float(q)
        except:
            continue
            
        if c is not None and str(c).strip() != "":
            last_ctn = str(c).strip()
            
        raw_items.append({
            "ctn": last_ctn,
            "model": sm,
            "inch": float(inch) if inch and str(inch).replace('.', '', 1).isdigit() else 0.0,
            "qty": qty_num,
            "nw": float(nw) if nw and str(nw).replace('.', '', 1).isdigit() else 0.0,
            "gw": float(gw) if gw and str(gw).replace('.', '', 1).isdigit() else 0.0,
            "brand": cur_brd
        })
        
    # 2. Group consecutive identical (brand, model, qty)
    n_raw = len(raw_items)
    groups = []
    i = 0
    while i < n_raw:
        j = i + 1
        while j < n_raw and raw_items[j]["model"] == raw_items[i]["model"] and raw_items[j]["qty"] == raw_items[i]["qty"] and raw_items[j]["brand"] == raw_items[i]["brand"]:
            j += 1
        cnt = j - i
        start_c = raw_items[i]["ctn"]
        end_c = raw_items[j-1]["ctn"]
        c_str = start_c if (cnt == 1 or start_c == end_c) else f"{start_c}-{end_c}"
        groups.append({
            "cno": c_str,
            "start_ctn": start_c,
            "end_ctn": end_c,
            "model": raw_items[i]["model"],
            "inch": raw_items[i]["inch"],
            "brand": raw_items[i]["brand"],
            "pcs": raw_items[i]["qty"],
            "ctns": cnt,
            "qty": raw_items[i]["qty"] * cnt,
            "nw": raw_items[i]["nw"],
            "gw": raw_items[i]["gw"],
            "tnw": round(raw_items[i]["nw"] * cnt, 3),
            "tgw": round(raw_items[i]["gw"] * cnt, 3)
        })
        i = j
        
    # 3. Model dictionaries
    dPCS_model = {}
    dCtns_model = {}
    dQty_model = {}
    
    dPCS_brand = {}
    dCtns_brand = {}
    dQty_brand = {}
    
    for g in groups:
        k_full = f"{g['brand']}_{normalize_model(g['model'])}"
        k_short = normalize_model(g['model'])
        
        if k_full not in dPCS_brand:
            dPCS_brand[k_full] = g["pcs"]
            dCtns_brand[k_full] = g["ctns"]
            dQty_brand[k_full] = g["qty"]
        else:
            dCtns_brand[k_full] += g["ctns"]
            dQty_brand[k_full] += g["qty"]
            
        if k_short not in dPCS_model:
            dPCS_model[k_short] = g["pcs"]
            dCtns_model[k_short] = g["ctns"]
            dQty_model[k_short] = g["qty"]
        else:
            dCtns_model[k_short] += g["ctns"]
            dQty_model[k_short] += g["qty"]

    unique_ctns = set(x["ctn"] for x in raw_items if x["ctn"] and str(x["ctn"]).strip().isdigit())
    grand_ctns = len(unique_ctns) if unique_ctns else len(set(x["ctn"] for x in raw_items if x["ctn"]))

    # 4. CI TAX
    ci_path = os.path.join("Templates", "TEMPLATE_CI.xlsx")
    if os.path.exists(ci_path):
        wb_ci = openpyxl.load_workbook(ci_path)
        ws_ci = wb_ci.active
        ws_ci["J3"] = f"INVOICE NO.: {inv_no}"
        ws_ci["J4"] = f"DATE: {inv_date}"
        
        prices = {}
        for r in range(13, ws_ci.max_row):
            m = str(ws_ci.cell(row=r, column=3).value or "").strip()
            p = str(ws_ci.cell(row=r, column=9).value or "").strip()
            desc = str(ws_ci.cell(row=r, column=2).value or "").strip()
            b = extract_brand(desc)
            brd_name = "TAKA PRO" if b == "TAKAPRO" else (b or "TAKA PRO")
            if m and m not in ["NAN", "NONE", ""] and not m.startswith("="):
                norm_m = normalize_model(m)
                try:
                    price_val = float(p)
                    prices[f"{brd_name}_{norm_m}"] = price_val
                    if norm_m not in prices:
                        prices[norm_m] = price_val
                except:
                    pass

        rows_to_delete_ci = []
        grand_qty_ci = 0
        grand_amt_ci = 0
        total_row_idx = -1
        
        for r in range(13, ws_ci.max_row + 1):
            col_a_val = str(ws_ci.cell(row=r, column=1).value or "").strip().upper()
            if col_a_val == "TOTAL":
                total_row_idx = r
                break
            
            if col_a_val.isdigit():
                m = str(ws_ci.cell(row=r, column=3).value or "").strip()
                desc = str(ws_ci.cell(row=r, column=2).value or "").strip()
                b = extract_brand(desc)
                brd_name = "TAKA PRO" if b == "TAKAPRO" else (b or "TAKA PRO")
                norm_m = normalize_model(m)
                key_full = f"{brd_name}_{norm_m}"
                
                qty = dQty_brand.get(key_full, dQty_model.get(norm_m, 0))
                pcs = dPCS_model.get(norm_m, dPCS_brand.get(key_full, 0))
                ctns = dCtns_model.get(norm_m, dCtns_brand.get(key_full, 0))
                
                if qty == 0:
                    rows_to_delete_ci.append(r)
                else:
                    price = prices.get(key_full, prices.get(norm_m, 0))
                    amt = round(qty * price, 2)
                    
                    ws_ci.cell(row=r, column=6).value = qty
                    ws_ci.cell(row=r, column=7).value = int(pcs) if float(pcs).is_integer() else pcs
                    ws_ci.cell(row=r, column=8).value = int(ctns)
                    ws_ci.cell(row=r, column=9).value = price
                    ws_ci.cell(row=r, column=10).value = amt
                    
                    grand_qty_ci += qty
                    grand_amt_ci += amt

        for r in reversed(rows_to_delete_ci):
            ws_ci.delete_rows(r, 1)
            if total_row_idx != -1: total_row_idx -= 1
            
        if total_row_idx != -1:
            ws_ci.cell(row=total_row_idx, column=6).value = grand_qty_ci
            ws_ci.cell(row=total_row_idx, column=8).value = grand_ctns
            ws_ci.cell(row=total_row_idx, column=10).value = round(grand_amt_ci, 2)
            
            for check_r in range(total_row_idx, total_row_idx + 5):
                val = str(ws_ci.cell(row=check_r, column=1).value or "").strip()
                if "SAY" in val:
                    ws_ci.cell(row=check_r, column=1).value = "SAY IN WORD: " + float_to_currency_words(grand_amt_ci)
                    break
                    
        item_idx = 1
        for r in range(13, ws_ci.max_row + 1):
            val = str(ws_ci.cell(row=r, column=1).value or "").strip()
            if val == "TOTAL": break
            if val.isdigit() or val != "None":
                if ws_ci.cell(row=r, column=3).value:
                    ws_ci.cell(row=r, column=1).value = item_idx
                    item_idx += 1

        ci_bytes = io.BytesIO()
        wb_ci.save(ci_bytes)
        zip_file.writestr(f"{inv_no}_COMMERCIAL_INVOICE_TAX.xlsx", ci_bytes.getvalue())

    # 5. PL TAX
    pl_path = os.path.join("Templates", "TEMPLATE_PL.xlsx")
    if os.path.exists(pl_path):
        wb_pl = openpyxl.load_workbook(pl_path)
        ws_pl = wb_pl.active
        ws_pl["K3"] = f"INVOICE NO.: {inv_no}"
        ws_pl["K4"] = f"DATE: {inv_date}"
        
        ws_pl.delete_rows(11, ws_pl.max_row - 10)
        
        ranges_to_remove = [rng for rng in ws_pl.merged_cells.ranges if rng.min_row >= 11]
        for rng in ranges_to_remove:
            ws_pl.merged_cells.remove(rng)
            
        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        center_align = Alignment(horizontal="center", vertical="center")
        left_align = Alignment(horizontal="left", vertical="center")
        
        row_idx = 11
        current_brand = None
        prev_cno = None
        
        grand_qty_pl = 0
        grand_nw_pl = 0
        grand_gw_pl = 0
        
        for g in groups:
            if g["brand"] != current_brand:
                current_brand = g["brand"]
                display_brand = current_brand
                ws_pl.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=11)
                cell = ws_pl.cell(row=row_idx, column=1)
                cell.value = f"RUBBER BELT (V-RIBBED) {display_brand}"
                cell.font = Font(bold=True)
                cell.alignment = left_align
                for c in range(1, 12):
                    ws_pl.cell(row=row_idx, column=c).border = thin_border
                row_idx += 1
                prev_cno = None

            is_same_carton = (prev_cno is not None and g["cno"] == prev_cno)
            pcs_ctn = int(g["pcs"]) if float(g["pcs"]).is_integer() else round(g["pcs"], 2)
            
            ws_pl.cell(row=row_idx, column=1).value = None if is_same_carton else g["cno"]
            ws_pl.cell(row=row_idx, column=2).value = f"Rubber Belt (V-Ribbed) {current_brand} {g['model']}"
            ws_pl.cell(row=row_idx, column=3).value = g["model"]
            ws_pl.cell(row=row_idx, column=4).value = int(g["inch"]) if float(g["inch"]).is_integer() else g["inch"]
            ws_pl.cell(row=row_idx, column=5).value = g["qty"]
            ws_pl.cell(row=row_idx, column=6).value = pcs_ctn
            
            if not is_same_carton:
                ws_pl.cell(row=row_idx, column=7).value = g["ctns"]
                if g["nw"] > 0: ws_pl.cell(row=row_idx, column=8).value = g["nw"]
                if g["gw"] > 0: ws_pl.cell(row=row_idx, column=9).value = g["gw"]
                if g["tnw"] > 0: ws_pl.cell(row=row_idx, column=10).value = round(g["tnw"], 2)
                if g["tgw"] > 0: ws_pl.cell(row=row_idx, column=11).value = round(g["tgw"], 2)
                grand_nw_pl += g["tnw"]
                grand_gw_pl += g["tgw"]
            else:
                ws_pl.cell(row=row_idx, column=7).value = None
                ws_pl.cell(row=row_idx, column=8).value = None
                ws_pl.cell(row=row_idx, column=9).value = None
                ws_pl.cell(row=row_idx, column=10).value = None
                ws_pl.cell(row=row_idx, column=11).value = None
                
            grand_qty_pl += g["qty"]
            
            for c in range(1, 12):
                cell = ws_pl.cell(row=row_idx, column=c)
                cell.border = thin_border
                cell.alignment = center_align
                
            prev_cno = g["cno"]
            row_idx += 1

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
        
        ws_pl.cell(row=row_idx, column=5).value = grand_qty_pl
        ws_pl.cell(row=row_idx, column=7).value = grand_ctns
        ws_pl.cell(row=row_idx, column=10).value = round(grand_nw_pl, 0)
        ws_pl.cell(row=row_idx, column=11).value = round(grand_gw_pl, 0)
        
        for c in range(1, 12):
            cell = ws_pl.cell(row=row_idx, column=c)
            cell.border = thin_border
            cell.fill = fill
            cell.font = Font(bold=True, color="FFFFFF")
            
        row_idx += 2
        ws_pl.cell(row=row_idx, column=1).value = "SHIPPING MARKS:"
        ws_pl.cell(row=row_idx, column=3).value = "TAKA PRO"
        ws_pl.cell(row=row_idx+1, column=3).value = "TAKA"
        ws_pl.cell(row=row_idx+3, column=1).value = f"TOTAL: {grand_qty_pl:,.0f} PCS / {grand_ctns} CARTONS"
        ws_pl.cell(row=row_idx+4, column=1).value = f"G.W.: {round(grand_gw_pl, 3):,.3f} KGS"
        ws_pl.cell(row=row_idx+5, column=1).value = f"N.W.: {round(grand_nw_pl, 3):,.3f} KGS"
        ws_pl.cell(row=row_idx+6, column=1).value = "MEAS.: __________ CBM"
        
        pl_bytes = io.BytesIO()
        wb_pl.save(pl_bytes)
        zip_file.writestr(f"{inv_no}_PACKING_LIST_TAX.xlsx", pl_bytes.getvalue())

    # 6. DRAFT FE
    fe_path = os.path.join("Templates", "TEMPLATE_FE.xlsx")
    if os.path.exists(fe_path):
        wb_fe = openpyxl.load_workbook(fe_path)
        ws_fe = wb_fe.active
        
        for r in range(23, 42):
            for c in range(1, 10):
                cell = ws_fe.cell(row=r, column=c)
                if type(cell).__name__ != 'MergedCell':
                    cell.value = None

        fe_groups_dict = {}
        for g in groups:
            hs, sdesc = get_hs_info(g["inch"])
            brd = g["brand"]
            k = (brd, hs, sdesc)
            if k not in fe_groups_dict:
                fe_groups_dict[k] = {"ctns": 0, "qty": 0, "prefixes": set()}
            fe_groups_dict[k]["ctns"] += g["ctns"]
            fe_groups_dict[k]["qty"] += g["qty"]
            fe_groups_dict[k]["prefixes"].add(get_prefix(g["model"]))
            
        brand_priority = {"TAKA PRO": 1, "TAKA": 2, "YAMATACHI": 3}
        sorted_fe_keys = sorted(
            fe_groups_dict.keys(),
            key=lambda k: (brand_priority.get(k[0], 99), k[1], k[2])
        )
        
        row_idx = 23
        item_num = 1
        total_fe_qty = 0
        
        for k in sorted_fe_keys:
            brd, hs_code, size_desc = k
            g_info = fe_groups_dict[k]
            total_cartons = int(g_info["ctns"])
            total_qty = int(g_info["qty"])
            total_fe_qty += total_qty
            prefixes = ", ".join(sorted(g_info["prefixes"]))
            
            ctn_words = number_to_words(total_cartons)
            desc = f"{ctn_words} ({total_cartons}) CTNS OF RUBBER BELT (V-RIBBED) {brd} {prefixes} {size_desc}"
            
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

    # 7. CONTRACT
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
        grand_ctns_contract = 0
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
                row.cells[6].text = f" {grand_ctns_contract:,.0f} "
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
                brd_name = "TAKA PRO" if b == "TAKAPRO" else (b or "TAKA PRO")
                m = row.cells[3].text.strip()
                norm_m = normalize_model(m)
                key_full = f"{brd_name}_{norm_m}"
                
                qty = dQty_brand.get(key_full, dQty_model.get(norm_m, 0))
                pcs = dPCS_brand.get(key_full, dPCS_model.get(norm_m, 0))
                ctns = dCtns_brand.get(key_full, dCtns_model.get(norm_m, 0))
                
                if qty == 0:
                    rows_to_delete.append(row)
                else:
                    price = prices.get(key_full, prices.get(norm_m))
                    if price is None:
                        price_str = row.cells[7].text
                        price = float(re.sub(r'[^\d.]', '', price_str))
                    amt = round(qty * price, 2)
                    
                    row.cells[4].text = f" {qty:,.0f} "
                    row.cells[5].text = f" {int(pcs) if float(pcs).is_integer() else pcs} "
                    row.cells[6].text = f" {ctns:,.0f} "
                    row.cells[7].text = f" $ {price:,.2f} "
                    row.cells[8].text = f" $ {amt:,.2f} "
                    
                    accum_qty += qty
                    accum_ctns += ctns
                    grand_qty += qty
                    grand_ctns_contract += ctns
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


# ----------------- GIAO DIỆN CHÍNH (MAIN UI) -----------------
st.title("📦 Tự Động Hóa Chứng Từ Nhập Khẩu")
st.markdown("Hệ thống tự động hóa lập bộ chứng từ Hải quan: **Commercial Invoice, Packing List, Hợp đồng Contract, Form E**.")
st.markdown("---")

# Sidebar
st.sidebar.title("⚙️ Cấu Hình Chung")
sidebar_rate = st.sidebar.number_input("Tỷ giá RMB sang USD:", value=6.7400, format="%.4f", step=0.01)
st.sidebar.markdown("---")
st.sidebar.markdown("### 💡 Hướng dẫn nhanh")
st.sidebar.markdown("""
1. **Tải file:** Kéo thả 1 hoặc 2 file Excel chứng từ NCC vào ô bên phải.
2. **Kiểm tra:** Hệ thống tự nhận diện file Packing List, Invoice và nhà cung cấp.
3. **Xử lý:** Bấm nút **Bắt đầu xử lý**.
4. **Tải về:** Bấm nút tải file ZIP hoàn chỉnh.
""")

if st.sidebar.button("🔄 Làm mới / Tải file khác"):
    for k in ["processed_zip", "processed_filename", "processed_msg"]:
        if k in st.session_state:
            del st.session_state[k]
    st.rerun()

# 1. Main File Uploader
st.subheader("1️⃣ Tải lên file chứng từ NCC (.xlsx)")
uploaded_files = st.file_uploader(
    "Kéo thả 1 file Excel (có 2 sheet) HOẶC kéo thả cả 2 file (Packing List + Invoice) vào đây:",
    type=["xlsx"],
    accept_multiple_files=True,
    help="Hỗ trợ cả trường hợp 1 file Excel chứa cả 2 sheet hoặc 2 file Excel riêng biệt."
)

if not uploaded_files:
    st.info("""
    👉 **Chưa có file nào được tải lên.**
    
    Vui lòng kéo thả file Excel chứng từ của Nhà cung cấp vào ô bên trên để bắt đầu!
    - Hỗ trợ định dạng chứng từ của **Taizhou Norton** và **Feilizhou**.
    - Hệ thống sẽ tự động đối soát, phân bổ thùng, áp mã HS và tạo trọn bộ chứng từ chuẩn.
    """)
    if "processed_zip" in st.session_state:
        st.markdown("---")
        st.success(st.session_state.get("processed_msg", "Đã xử lý thành công!"))
        st.download_button(
            label="📥 TẢI VỀ BỘ CHỨNG TỪ ĐÃ TẠO GẦN NHẤT (ZIP)",
            data=st.session_state["processed_zip"],
            file_name=st.session_state.get("processed_filename", "CHUNGTU.zip"),
            mime="application/zip",
            type="primary",
            use_container_width=True
        )
else:
    df_pl_raw, df_inv_raw, norton_pl_file, pl_name, inv_name, err = load_excel_files(uploaded_files)
    if err:
        st.error(f"❌ {err}")
    else:
        col_a, col_b = st.columns(2)
        col_a.success(f"📋 **Packing List đã nhận:** {pl_name}")
        col_b.success(f"💵 **Invoice đã nhận:** {inv_name}")
        
        # Auto detect supplier
        full_inv_text = " ".join([str(x).upper() for row in df_inv_raw.values for x in row if pd.notna(x)])
        is_norton = ("NORTON" in full_inv_text) or ("TAIZHOU" in full_inv_text)
        
        st.subheader("2️⃣ Xác nhận cấu hình & Nhà cung cấp")
        c_sup, c_rate = st.columns(2)
        with c_sup:
            supplier = st.radio(
                "Nhà cung cấp:",
                ["Taizhou Norton", "Feilizhou"],
                index=0 if is_norton else 1,
                horizontal=True
            )
        with c_rate:
            exchange_rate = st.number_input("Tỷ giá RMB sang USD:", value=sidebar_rate, format="%.4f", key="rate_main")
            
        st.markdown("---")
        if st.button("🚀 BẮT ĐẦU XỬ LÝ & TẠO CHỨNG TỪ", type="primary", use_container_width=True):
            with st.spinner(f"Đang tự động tính toán phân bổ thùng và tạo bộ chứng từ cho {supplier}..."):
                try:
                    inv_no, inv_date = "UNKNOWN", "UNKNOWN"
                    for r in range(min(15, len(df_inv_raw))):
                        for c in range(min(10, df_inv_raw.shape[1])):
                            val = str(df_inv_raw.iloc[r, c]).strip().upper()
                            if "INVOICE NO" in val:
                                parts = str(df_inv_raw.iloc[r, c]).split(":")
                                if len(parts) > 1 and parts[-1].strip():
                                    inv_no = parts[-1].strip()
                                elif c+1 < df_inv_raw.shape[1]:
                                    inv_no = str(df_inv_raw.iloc[r, c+1]).strip()
                            if "DATE" in val:
                                parts = str(df_inv_raw.iloc[r, c]).split(":")
                                if len(parts) > 1 and parts[-1].strip():
                                    inv_date = parts[-1].strip()
                                elif c+1 < df_inv_raw.shape[1]:
                                    inv_date = str(df_inv_raw.iloc[r, c+1]).strip()
                    
                    if inv_no.upper() in ["INVOICE NO.", "INVOICE NO", "INVOICE NO:", ""] or "NAN" in inv_no.upper():
                        inv_no = "FLZ_INV"
                    if inv_date.upper() in ["DATE:", "DATE", ""] or "NAN" in inv_date.upper():
                        inv_date = "TODAY"
                        
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        if supplier == "Taizhou Norton":
                            new_supplier_logic.process(
                                zip_file, df_pl_raw, df_inv_raw, norton_pl_file,
                                inv_no, inv_date, exchange_rate,
                                extract_brand, normalize_model, get_prefix, get_hs_info,
                                None, None, float_to_currency_words, number_to_words
                            )
                        elif supplier == "Feilizhou":
                            process_feilizhou(
                                zip_file, df_pl_raw, df_inv_raw, inv_no, inv_date, exchange_rate
                            )
                            
                    st.session_state["processed_zip"] = zip_buffer.getvalue()
                    st.session_state["processed_filename"] = f"CHUNGTU_{inv_no}.zip"
                    st.session_state["processed_msg"] = f"✅ Đã xử lý phân bổ thùng và tạo trọn bộ chứng từ thành công! Số hóa đơn: {inv_no}."
                    st.rerun()
                except Exception as e:
                    import traceback
                    st.error(f"Lỗi xử lý: {e}")
                    st.code(traceback.format_exc())

        if "processed_zip" in st.session_state:
            st.markdown("---")
            st.success(st.session_state["processed_msg"])
            st.download_button(
                label="📥 TẢI VỀ TRỌN BỘ 4 CHỨNG TỪ (ZIP)",
                data=st.session_state["processed_zip"],
                file_name=st.session_state["processed_filename"],
                mime="application/zip",
                type="primary",
                use_container_width=True
            )
