import pandas as pd
import openpyxl
import os
import io
import re
from docx import Document
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

def process(zip_file, df_pl_raw, df_inv_raw, norton_pl_file, inv_no, inv_date, exchange_rate, extract_brand, normalize_model, get_prefix, get_hs_info, model_cartons, pl_groups, float_to_currency_words, number_to_words):
    # 1. Read RMB prices and sizes from input CI
    in_prices = {}
    in_sizes = {}
    for idx, row in df_inv_raw.iterrows():
        try:
            m = str(row[3]).strip()
            p = str(row[7]).strip()
            s = str(row[2]).strip()
            desc = str(row[1]).strip()
            if m and m not in ["NAN", "NONE", ""]:
                b = extract_brand(desc)
                if b:
                    key = f"{b}_{normalize_model(m)}"
                    if p and p not in ["NAN", "NONE", ""]:
                        price_val = float(re.sub(r'[^\d.]', '', p))
                        in_prices[key] = price_val
                    if s and s not in ["NAN", "NONE", ""]:
                        size_val = float(re.sub(r'[^\d.]', '', s))
                        in_sizes[key] = size_val / 2.54 # Convert CM to INCH for our logic
        except Exception:
            pass

    # 2. Build items flat list for FE and groupings
    items_flat = []
    current_brand = "UNKNOWN"
    current_ctns = []
    assigned_ctns = set()
    for idx, row in df_pl_raw.iterrows():
        row_str = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
        b = extract_brand(row_str)
        if b: current_brand = b

        col_a = str(row[0]).strip().upper()
        if "RUBBER" in col_a:
            assigned_ctns.clear()
            
        col_b = str(row[2]).strip().upper() # In Norton PL, Model is in col 2 (BELT NO.)
        
        if pd.isna(row[2]) or col_b in ["NAN", "NONE", ""]: continue
        if "BELT NO" in col_b or "RUBBER" in col_b: continue
        
        if re.match(r"^\d+$", col_a): current_ctns = [int(col_a)]
        elif re.match(r"^\d+-\d+$", col_a):
            parts = col_a.split("-")
            start = int(parts[0])
            end = int(parts[1])
            current_ctns = list(range(start, end + 1))
            
        ctn_count = 0
        for c in current_ctns:
            if c not in assigned_ctns:
                assigned_ctns.add(c)
                ctn_count += 1
                
        qty = float(row[3]) if pd.notna(row[3]) else 0
        pcs_ctn = float(row[4]) if pd.notna(row[4]) else 0
        
        m_normalized = normalize_model(col_b)
        key = f"{current_brand}_{m_normalized}"
        
        inch_val = in_sizes.get(key, 0)
        hs_code, size_desc = get_hs_info(inch_val)
        
        items_flat.append({
            "brand": current_brand,
            "model": m_normalized,
            "qty": qty,
            "pcs_ctn": pcs_ctn,
            "assigned_ctns": ctn_count,
            "inch": inch_val,
            "hs_code": hs_code,
            "size_desc": size_desc
        })

    df_items = pd.DataFrame(items_flat)
    if len(df_items) > 0:
        df_items["prefix"] = df_items["model"].apply(get_prefix)
        df_items[["hs_code", "size_desc"]] = df_items.apply(lambda r: pd.Series(get_hs_info(r["inch"])), axis=1)

        df_items[["hs_code", "size_desc"]] = df_items.apply(lambda r: pd.Series(get_hs_info(r["inch"])), axis=1)


        brand_order = {"TAKAPRO": 1, "TAKA": 2, "YAMATACHI": 3}
        df_items['brand_order'] = df_items['brand'].map(brand_order).fillna(99)
        
        fe_groups = df_items.groupby(["brand_order", "brand", "hs_code", "size_desc"])

    # 3. Form E Generation
    fe_path = os.path.join("Templates", "TEMPLATE_FE_NORTON.xlsx")
    if os.path.exists(fe_path) and len(df_items) > 0:
        wb_fe = openpyxl.load_workbook(fe_path)
        ws_fe = wb_fe.active
        
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
            desc = f"{ctn_words} ({total_cartons}) CTNS OF RUBBER V BELT {display_brand} {prefixes} {size_desc}"
            
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
            
            if row_idx > 30:
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

    # 4. CI Generation
    ci_path = os.path.join("Templates", "TEMPLATE_CI_NORTON.xlsx")
    if os.path.exists(ci_path) and len(df_items) > 0:
        wb_ci = openpyxl.load_workbook(ci_path)
        ws_ci = wb_ci.active
        ws_ci["H7"] = f"INV. NO. {inv_no}"
        ws_ci["H8"] = f"DATE: {inv_date}"

        row_idx = 13
        item_idx = 1
        grand_qty_ci = 0
        grand_ctns_ci = 0
        grand_amt_ci = 0

        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        bold_font = Font(bold=True)
        center_align = Alignment(horizontal="center", vertical="center")
        
        for name, group in fe_groups:
            _, brand, hs_code, size_desc = name
            prefixes = ", ".join(sorted(group["prefix"].unique()))
            display_brand = "TAKA PRO" if brand == "TAKAPRO" else brand
            
            # Group header
            header_desc = f"Rubber V Belt {display_brand} {prefixes} {size_desc}"
            ws_ci.cell(row=row_idx, column=1).value = header_desc
            ws_ci.cell(row=row_idx, column=1).font = bold_font
            ws_ci.cell(row=row_idx, column=5).value = int(group["qty"].sum())
            ws_ci.cell(row=row_idx, column=7).value = int(group["assigned_ctns"].sum())
            for c in range(1, 15):
                ws_ci.cell(row=row_idx, column=c).border = thin_border
                if c not in [1, 5, 7]:
                    ws_ci.cell(row=row_idx, column=c).value = ""
            row_idx += 1
            
            # Group items
            item_groups = group.groupby(["model"])
            for m, m_group in item_groups:
                m_qty = m_group["qty"].sum()
                m_ctns = m_group["assigned_ctns"].sum()
                
                key = f"{brand}_{normalize_model(m[0])}"
                rmb_price = in_prices.get(key, 0)
                usd_price = round(rmb_price / exchange_rate, 2)
                amt = m_qty * usd_price
                
                ws_ci.cell(row=row_idx, column=1).value = item_idx
                ws_ci.cell(row=row_idx, column=1).alignment = center_align
                ws_ci.cell(row=row_idx, column=1).font = Font(bold=False)
                ws_ci.cell(row=row_idx, column=2).value = f"Rubber V Belt {display_brand} {m[0]}"
                ws_ci.cell(row=row_idx, column=3).value = float(m_group.iloc[0]["inch"]) * 2.54 # Size in CM
                ws_ci.cell(row=row_idx, column=4).value = m[0]
                ws_ci.cell(row=row_idx, column=5).value = m_qty
                
                pcs_ctn_val = m_group.iloc[0]["pcs_ctn"] if m_group.iloc[0]["pcs_ctn"] > 0 else (m_qty/m_ctns if m_ctns > 0 else 0)
                ws_ci.cell(row=row_idx, column=6).value = pcs_ctn_val
                
                ws_ci.cell(row=row_idx, column=7).value = m_ctns
                ws_ci.cell(row=row_idx, column=8).value = usd_price
                ws_ci.cell(row=row_idx, column=9).value = round(amt, 2)
                
                brand_code = "YAMA" if brand == "YAMATACHI" else ("TA" if brand == "TAKA" else brand)
                prefix_code = m_group.iloc[0]["prefix"] + "T"
                model_num = m[0].replace(m_group.iloc[0]["prefix"], "")
                padded_num = model_num.zfill(3) if model_num.isdigit() else model_num
                ws_ci.cell(row=row_idx, column=10).value = f"CR.{prefix_code}.{padded_num}.{brand_code}"
                ws_ci.cell(row=row_idx, column=11).value = f"{display_brand} trơn {m[0]}"
                ws_ci.cell(row=row_idx, column=12).value = ""
                ws_ci.cell(row=row_idx, column=13).value = ""
                ws_ci.cell(row=row_idx, column=14).value = hs_code
                
                for c in range(1, 15):
                    ws_ci.cell(row=row_idx, column=c).border = thin_border
                
                grand_qty_ci += m_qty
                grand_ctns_ci += m_ctns
                grand_amt_ci += round(amt, 2)
                
                item_idx += 1
                row_idx += 1
                
        # Write TOTAL row
        total_row = row_idx
        ws_ci.cell(row=total_row, column=1).value = "TOTAL"
        ws_ci.cell(row=total_row, column=1).font = bold_font
        for c in range(2, 15):
            ws_ci.cell(row=total_row, column=c).value = ""
        ws_ci.cell(row=total_row, column=5).value = grand_qty_ci
        ws_ci.cell(row=total_row, column=5).font = bold_font
        ws_ci.cell(row=total_row, column=7).value = grand_ctns_ci
        ws_ci.cell(row=total_row, column=7).font = bold_font
        ws_ci.cell(row=total_row, column=9).value = round(grand_amt_ci, 2)
        ws_ci.cell(row=total_row, column=9).font = bold_font
        
        for c in range(1, 15):
            ws_ci.cell(row=total_row, column=c).border = thin_border
            ws_ci.cell(row=total_row, column=c).alignment = Alignment(horizontal="center")
            
        say_row = total_row + 1
        ws_ci.cell(row=say_row, column=1).value = "SAY IN WORD: " + float_to_currency_words(grand_amt_ci)
        ws_ci.cell(row=say_row, column=1).font = bold_font
        for c in range(1, 15):
            ws_ci.cell(row=say_row, column=c).border = Border() # no border
        ws_ci.merge_cells(start_row=say_row, start_column=1, end_row=say_row, end_column=14)



        ci_bytes = io.BytesIO()
        wb_ci.save(ci_bytes)
        zip_file.writestr(f"{inv_no}_COMMERCIAL_INVOICE_TAX.xlsx", ci_bytes.getvalue())

    # 5. Contract Generation
    contract_path = os.path.join("Templates", "TEMPLATE_CONTRACT_NORTON.docx")
    if os.path.exists(contract_path) and len(df_items) > 0:
        doc = Document(contract_path)
        
        for p in doc.paragraphs:
            if p.text.startswith("No:"): p.text = f"No: {inv_no}"
            elif p.text.startswith("Date:"): p.text = f"Date: {inv_date}"
            elif "Total Value of Contract:" in p.text:
                parts = p.text.split(":")
                p.text = f"{parts[0]}: {grand_amt_ci:,.2f} USD"
            elif "Payment 100% total Value of contract:" in p.text:
                p.text = f"Payment 100% total Value of contract: {grand_amt_ci:,.2f} USD by TT after shipment with the favor account of Seller:"
            
        table = doc.tables[-1]
        
        item_idx = 1
        for name, group in fe_groups:
            _, brand, hs_code, size_desc = name
            prefixes = ", ".join(sorted(group["prefix"].unique()))
            display_brand = "TAKA PRO" if brand == "TAKAPRO" else brand
            
            # Group header row
            row_cells = table.add_row().cells
            header_desc = f"Rubber V Belt {display_brand} {prefixes} {size_desc}"
            row_cells[0].text = header_desc
            row_cells[1].text = header_desc
            row_cells[2].text = header_desc
            row_cells[3].text = header_desc
            row_cells[4].text = f" {int(group['qty'].sum()):,.0f} "
            row_cells[5].text = ""
            row_cells[6].text = f" {int(group['assigned_ctns'].sum()):,.0f} "
            row_cells[7].text = ""
            row_cells[8].text = ""
            
            # Group items
            item_groups = group.groupby(["model"])
            for m, m_group in item_groups:
                m_qty = m_group["qty"].sum()
                m_ctns = m_group["assigned_ctns"].sum()
                
                key = f"{brand}_{normalize_model(m[0])}"
                rmb_price = in_prices.get(key, 0)
                usd_price = round(rmb_price / exchange_rate, 2)
                amt = m_qty * usd_price
                
                row_cells = table.add_row().cells
                row_cells[0].text = str(item_idx)
                row_cells[1].text = f"Rubber V Belt {display_brand} {m[0]}"
                row_cells[2].text = str(round(float(m_group.iloc[0]["inch"]) * 2.54, 1))
                row_cells[3].text = m[0]
                row_cells[4].text = f" {m_qty:,.0f} "
                
                pcs_ctn_val = m_group.iloc[0]["pcs_ctn"] if m_group.iloc[0]["pcs_ctn"] > 0 else (m_qty/m_ctns if m_ctns > 0 else 0)
                row_cells[5].text = f" {int(pcs_ctn_val) if float(pcs_ctn_val).is_integer() else round(pcs_ctn_val, 2)} "
                row_cells[6].text = f" {m_ctns:,.0f} "
                row_cells[7].text = f" $ {usd_price:,.2f} "
                row_cells[8].text = f" $ {round(amt, 2):,.2f} "
                
                item_idx += 1
                
        # Total row
        row_cells = table.add_row().cells
        row_cells[0].text = "TOTAL"
        row_cells[4].text = f" {grand_qty_ci:,.0f} "
        row_cells[6].text = f" {grand_ctns_ci:,.0f} "
        row_cells[8].text = f" $ {grand_amt_ci:,.2f} "
        
        say_row = table.add_row().cells
        for i in range(9):
            say_row[i].text = "SAY IN WORD: " + float_to_currency_words(grand_amt_ci)
            
        contract_bytes = io.BytesIO()
        doc.save(contract_bytes)
        zip_file.writestr(f"CONTRACT_{inv_no}.docx", contract_bytes.getvalue())

    # 6. Norton PL (Pass-through)
    
    if hasattr(norton_pl_file, "getvalue"):
        zip_file.writestr(f"{inv_no}_PACKING_LIST_TAX.xlsx", norton_pl_file.getvalue())
    else:
        norton_pl_file.seek(0)
        zip_file.writestr(f"{inv_no}_PACKING_LIST_TAX.xlsx", norton_pl_file.read())


def get_pl_bytes(df_pl_raw):
    # Simply save the original dataframe to an excel file, to pass it through exactly as is
    # Wait, df_pl_raw has header=None, we can just save it.
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_pl_raw.to_excel(writer, index=False, header=False)
    return output.getvalue()
