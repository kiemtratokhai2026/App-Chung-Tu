import sys
import pandas as pd
import mapping_utils
import openpyxl
import re
import math
from datetime import datetime
from openpyxl.styles import Font, Alignment, Border, Side
import zipfile
import io
import os
from copy import copy
from docx import Document

def process(zip_file, df_pl_raw, df_inv_raw, pl_filename, invoice_no, invoice_date, exchange_rate, extract_brand, normalize_model, get_prefix, get_hs_info, contract_template_path, fe_template_path, float_to_currency_words, number_to_words):
    
    # 1. Parse PL for assigned_ctns
    assigned_ctns = set()
    model_cartons = {}
    
    for idx, row in df_pl_raw.iterrows():
        col_a = str(row[0]).strip().upper()
        col_b = str(row[2]).strip().upper()
        
        if pd.isna(row[2]) or col_b in ["NAN", "NONE", ""]: continue
        if "BELT NO" in col_b or "RUBBER" in col_b: continue
        
        row_str = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
        pl_brand = extract_brand(row_str)
        if pl_brand is None: pl_brand = ""
        m_normalized = normalize_model(col_b)
        key = (m_normalized, pl_brand)
        
        try:
            ctn_count = float(row[5])
            if pd.isna(ctn_count): ctn_count = 0
        except:
            ctn_count = 0
            
        if ctn_count >= 0:
            if key not in model_cartons:
                model_cartons[key] = 0
            model_cartons[key] += ctn_count

    # 2. Parse INPUT CI directly for exact groups and items
    groups = []
    current_group = None
    seen_item = False
    
    for idx, row in df_inv_raw.iterrows():
        if pd.isna(row[0]) and pd.isna(row[1]): continue
        
        c0 = str(row[0]).strip().upper()
        if c0 == 'NO.':
            seen_item = True
            continue
            
        if seen_item and pd.notna(row[0]) and pd.isna(row[1]) and pd.isna(row[2]):
            if "TOTAL" in c0: continue
            header_text = str(row[0]).strip()
            current_group = {"header": header_text, "items": []}
            groups.append(current_group)
    groups = []
    current_group = None
    seen_item = False
    
    for idx, row in df_inv_raw.iterrows():
        if pd.isna(row[0]) and pd.isna(row[1]): continue
        
        c0 = str(row[0]).strip().upper()
        if c0 == 'NO.':
            seen_item = True
            continue
            
        if seen_item and pd.notna(row[0]) and pd.isna(row[1]) and pd.isna(row[2]):
            if "TOTAL" in c0: continue
            header_text = str(row[0]).strip()
            current_group = {"header": header_text, "items": []}
            groups.append(current_group)
            
        elif seen_item and pd.notna(row[1]) and "TOTAL" not in str(row[1]).upper() and "RUBBER" in str(row[1]).upper():
            if current_group is None:
                current_group = {"header": "GENERAL", "items": []}
                groups.append(current_group)
                
            desc = str(row[1]).strip() if pd.notna(row[1]) else ""
            desc_upper = desc.upper()
            idx_rubber_desc = desc_upper.find("RUBBER")
            if idx_rubber_desc > 0:
                desc = desc[idx_rubber_desc:]

            belt_type = "Rubber V Belt"
            belt_type_fe = "RUBBER V BELT"
            if "V-RIBBED" in desc_upper:
                belt_type = "Rubber V-Ribbed Belt"
                belt_type_fe = "RUBBER BELT (V-RIBBED)"

            try: size = float(row[2])
            except: size = 0.0
            model = str(row[3]).strip() if pd.notna(row[3]) else ""
            try: qty = float(row[4])
            except: qty = 0.0
            try: pcs_ctn = float(row[5])
            except: pcs_ctn = 0.0
            
            try:
                price_rmb = float(row[7]) if len(row) > 7 else 0.0
                if pd.isna(price_rmb): price_rmb = 0.0
            except: price_rmb = 0.0
            
            price_usd = price_rmb / exchange_rate if exchange_rate else price_rmb
            amount_usd = qty * price_usd
            
            brand = extract_brand(desc)
            if brand is None: brand = ""
            m_normalized = normalize_model(model)
            key = (m_normalized, brand)
            if key in model_cartons:
                ctns = model_cartons[key]
                model_cartons[key] = 0 # Only assign to the first occurrence
            else:
                try: fallback = float(row[6]) if len(row) > 6 else 0.0
                except: fallback = 0.0
                ctns = math.ceil(fallback)
                
            brand_prefix = "UNKNOWN"
            if "TAKA PRO" in desc_upper: brand_prefix = "TAKAPRO"
            elif "TAKA" in desc_upper: brand_prefix = "TA"
            elif "YAMATACHI" in desc_upper: brand_prefix = "YAMA"
            elif "MICHELIN" in desc_upper: brand_prefix = "MIC"
            

            inch_val = size / 2.54 if size > 0 else 0
            norm_model = normalize_model(model)
            prefix = get_prefix(norm_model)
            misa_code = norm_model

            if "YAMATACHI" in brand: misa_code = f"CR.AT.023.YM.{prefix}"
            elif "MICHELIN" in brand: misa_code = f"CR.AT.025.MIC.{prefix}"
            elif "TAKA PRO" in brand: misa_code = f"CR.BR.030.TAPRO.{prefix}"
            elif "TAKA" in brand: misa_code = f"CR.BR.027.TA.{prefix}"

            hs_code, size_desc = get_hs_info(inch_val)
            sz_cm = inch_val * 2.54

            mapping = mapping_utils.lookup_mapping(norm_model, brand)
            if mapping:
                misa_code = mapping['internal_code'] if mapping['internal_code'] else misa_code
                if mapping['size_cm']:
                    sz_cm = mapping['size_cm']
                hs_code = mapping['hs_code'] if mapping['hs_code'] else hs_code

            # Grouping Logic
            if sz_cm < 60: size_desc_group = "401039_GROUP"
            elif sz_cm <= 180: size_desc_group = "FROM 60CM TO 180CM"
            elif sz_cm <= 240: size_desc_group = "FROM 180CM TO 240CM"
            else: size_desc_group = "401039_GROUP"
                
            current_group["items"].append({
                "brand": brand,
                "size_desc_group": size_desc_group,
                "sz_cm": sz_cm,
                "belt_type": belt_type,
                "belt_type_fe": belt_type_fe,
                "prefix": prefix,
                "desc": desc,
                "size": size,
                "model": model,
                "qty": qty,
                "pcs_ctn": pcs_ctn,
                "ctns": ctns,
                "price_usd": price_usd,
                "amount_usd": amount_usd,
                "misa_code": misa_code,
                "hs_code": hs_code,
                "size_desc": size_desc
            })

    # Filter empty groups
    groups = [g for g in groups if len(g["items"]) > 0]

    # 2.5 Dynamically regroup everything by brand, hs_code, and size_desc
    all_items = []
    for g in groups:
        all_items.extend(g["items"])
        
    if all_items:
        df_all = pd.DataFrame(all_items)
        brand_order = {"TAKAPRO": 1, "TAKA": 2, "YAMATACHI": 3}
        df_all['brand_order'] = df_all['brand'].map(brand_order).fillna(99)
        ci_groups = df_all.groupby(["brand_order", "brand", "belt_type", "belt_type_fe", "hs_code", "size_desc_group"], dropna=False)
        
        new_groups = []
        for name, group_df in ci_groups:
            _, brand, belt_type, belt_type_fe, hs_code, size_desc_group = name
            prefixes = ", ".join(sorted([p for p in group_df["prefix"].unique() if pd.notna(p) and p]))
            if not prefixes: prefixes = ""
            
            has_under_60 = any(i["sz_cm"] < 60 for _, i in group_df.iterrows())
            has_over_240 = any(i["sz_cm"] > 240 for _, i in group_df.iterrows())
            
            if size_desc_group == "401039_GROUP":
                if has_under_60 and has_over_240:
                    actual_size_desc = "under 60cm and over 240cm"
                elif has_under_60:
                    actual_size_desc = "under 60cm"
                elif has_over_240:
                    actual_size_desc = "over 240cm"
                else:
                    actual_size_desc = ""
            else:
                actual_size_desc = str(size_desc_group).lower() if pd.notna(size_desc_group) else ""
            
            header_text = f"{belt_type} {brand} {prefixes} {actual_size_desc}".strip()
            
            # Deduplicate by (model, brand)
            grouped = {}
            for _, item in group_df.iterrows():
                item_dict = item.to_dict()
                item_dict["actual_size_desc"] = actual_size_desc
                k = (item_dict["model"], item_dict["brand"])
                if k not in grouped:
                    grouped[k] = item_dict
                else:
                    grouped[k]["qty"] += item_dict["qty"]
                    grouped[k]["ctns"] += item_dict["ctns"]
                    grouped[k]["amount_usd"] += item_dict["amount_usd"]
                    
            new_groups.append({
                "header": header_text,
                "items": list(grouped.values())
            })
            
        groups = new_groups

    # 3. Write CI
    
    ci_path = os.path.join("Templates", "TEMPLATE_CI_NORTON.xlsx")
    wb_ci = openpyxl.load_workbook(ci_path)
    ws_ci = wb_ci.active
    
    ws_ci.cell(row=3, column=1).value = "INVOICE NO.: " + invoice_no
    ws_ci.cell(row=4, column=1).value = "DATE: " + invoice_date
    
    bold_font = Font(name='Times New Roman', size=11, bold=True)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'), right=openpyxl.styles.Side(style='thin'), top=openpyxl.styles.Side(style='thin'), bottom=openpyxl.styles.Side(style='thin'))

    row_idx = 13
    grand_qty_ci = 0
    grand_ctns_ci = 0
    grand_amt_ci = 0
    item_idx = 1
    
    for g in groups:
        ws_ci.cell(row=row_idx, column=1).value = g["header"]
        ws_ci.cell(row=row_idx, column=1).font = bold_font
        for c in range(1, 15):
            ws_ci.cell(row=row_idx, column=c).border = thin_border
            if c > 1:
                ws_ci.cell(row=row_idx, column=c).value = ""
        ws_ci.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=14)
        row_idx += 1
        
        for item in g["items"]:
            ws_ci.cell(row=row_idx, column=1).value = item_idx
            ws_ci.cell(row=row_idx, column=1).alignment = center_align
            ws_ci.cell(row=row_idx, column=1).font = Font(bold=False)
            ws_ci.cell(row=row_idx, column=2).value = item["desc"]
            ws_ci.cell(row=row_idx, column=3).value = item["size"]
            ws_ci.cell(row=row_idx, column=4).value = item["model"]
            ws_ci.cell(row=row_idx, column=5).value = item["qty"]
            ws_ci.cell(row=row_idx, column=6).value = item["pcs_ctn"]
            ws_ci.cell(row=row_idx, column=7).value = item["ctns"]
            ws_ci.cell(row=row_idx, column=8).value = round(item["price_usd"], 2)
            ws_ci.cell(row=row_idx, column=9).value = round(item["amount_usd"], 2)
            
            ws_ci.cell(row=row_idx, column=10).value = item["misa_code"]
            ws_ci.cell(row=row_idx, column=11).value = item["misa_code"].split(".")[-1] # or something similar
            ws_ci.cell(row=row_idx, column=12).value = ""
            ws_ci.cell(row=row_idx, column=13).value = ""
            ws_ci.cell(row=row_idx, column=14).value = item["hs_code"]
            
            for c in range(1, 15):
                ws_ci.cell(row=row_idx, column=c).border = thin_border
            
            grand_qty_ci += item["qty"]
            grand_ctns_ci += item["ctns"]
            grand_amt_ci += round(item["amount_usd"], 2)
            
            item_idx += 1
            row_idx += 1
            
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
        ws_ci.cell(row=say_row, column=c).border = openpyxl.styles.Border()
    ws_ci.merge_cells(start_row=say_row, start_column=1, end_row=say_row, end_column=14)

    ci_bytes = io.BytesIO()
    wb_ci.save(ci_bytes)
    zip_file.writestr(f"{invoice_no}_COMMERCIAL_INVOICE_TAX.xlsx", ci_bytes.getvalue())
    
    # Write Contract
    contract_path = os.path.join("Templates", "TEMPLATE_CONTRACT_NORTON.docx")
    if os.path.exists(contract_path) and len(groups) > 0:
        doc = Document(contract_path)
        for p in doc.paragraphs:
            if p.text.startswith("No:"): p.text = f"No: {invoice_no}"
            elif p.text.startswith("Date:"): p.text = f"Date: {invoice_date}"
            elif "Total Value of Contract:" in p.text:
                parts = p.text.split(":")
                p.text = f"{parts[0]}: {grand_amt_ci:,.2f} USD"
            elif "Payment 100% total Value of contract:" in p.text:
                p.text = f"Payment 100% total Value of contract: {grand_amt_ci:,.2f} USD by TT after shipment with the favor account of Seller:"
                
        table = doc.tables[-1]
        
        item_idx = 1
        for g in groups:
            row_cells = table.add_row().cells
            row_cells[0].text = g["header"]
            row_cells[1].text = g["header"]
            row_cells[2].text = g["header"]
            row_cells[3].text = g["header"]
            row_cells[4].text = f" {int(sum([i['qty'] for i in g['items']])):,.0f} "
            row_cells[5].text = ""
            row_cells[6].text = f" {int(sum([i['ctns'] for i in g['items']])):,.0f} "
            row_cells[7].text = ""
            row_cells[8].text = ""
            
            for item in g["items"]:
                row_cells = table.add_row().cells
                row_cells[0].text = str(item_idx)
                row_cells[1].text = item["desc"]
                row_cells[2].text = str(item["size"])
                row_cells[3].text = item["model"]
                row_cells[4].text = f" {item['qty']:,.0f} "
                pcs_ctn_val = item['pcs_ctn']
                row_cells[5].text = f" {int(pcs_ctn_val) if float(pcs_ctn_val).is_integer() else round(pcs_ctn_val, 2)} "
                row_cells[6].text = f" {item['ctns']:,.0f} "
                row_cells[7].text = f" $ {item['price_usd']:,.2f} "
                row_cells[8].text = f" $ {round(item['amount_usd'], 2):,.2f} "
                item_idx += 1
                
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
        zip_file.writestr(f"CONTRACT_{invoice_no}.docx", contract_bytes.getvalue())

    # --- FORM E GENERATION ---
    items_flat = []
    for g in groups:
        for item in g["items"]:
            items_flat.append(item)
    
    if len(items_flat) > 0:
        df_items = pd.DataFrame(items_flat)
        brand_order = {"TAKAPRO": 1, "TAKA": 2, "YAMATACHI": 3}
        df_items['brand_order'] = df_items['brand'].map(brand_order).fillna(99)
        fe_groups = df_items.groupby(["brand_order", "brand", "belt_type", "belt_type_fe", "hs_code", "size_desc_group"], dropna=False)

        fe_path = os.path.join("Templates", "TEMPLATE_FE_NORTON.xlsx")
        if os.path.exists(fe_path):
            from openpyxl.styles import Border, Side
            wb_fe = openpyxl.load_workbook(fe_path)
            ws_fe = wb_fe.active
            
            row_idx = 23
            item_num = 1
            total_fe_qty = 0
            
            for name, group in fe_groups:
                _, brand, belt_type, belt_type_fe, hs_code, size_desc_group = name
                prefixes = ", ".join(sorted([p for p in group["prefix"].unique() if pd.notna(p) and p]))
                if not prefixes: prefixes = ""
                
                actual_size_desc = group["actual_size_desc"].iloc[0]
                
                desc_text = f"{belt_type_fe}\n{brand}\n{prefixes}\n{actual_size_desc.upper()}"
                total_qty = group["qty"].sum()
                total_fe_qty += total_qty
                
                ws_fe.cell(row=row_idx, column=1).value = item_num
                ws_fe.cell(row=row_idx, column=7).value = desc_text
                ws_fe.cell(row=row_idx, column=8).value = f"{hs_code[:4]}.{hs_code[4:]}" if len(str(hs_code)) == 6 else hs_code
                ws_fe.cell(row=row_idx, column=9).value = "CTN"
                ws_fe.cell(row=row_idx, column=10).value = total_qty
                
                row_idx += 1
                item_num += 1

            ws_fe.cell(row=42, column=10).value = total_fe_qty
            
            thin_border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'), right=openpyxl.styles.Side(style='thin'), top=openpyxl.styles.Side(style='thin'), bottom=openpyxl.styles.Side(style='thin'))
            for r in range(23, row_idx):
                for c in range(1, 12):
                    ws_fe.cell(row=r, column=c).border = thin_border
                    
            fe_bytes = io.BytesIO()
            wb_fe.save(fe_bytes)
            zip_file.writestr(f"DRAFT FE_{invoice_no}_TAX.xlsx", fe_bytes.getvalue())

    return None
