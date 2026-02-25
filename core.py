# -*- coding: utf-8 -*-
"""
Created on Tue Feb 24 11:32:50 2026

@author: eivind
"""

import pandas as pd
import openpyxl
import os
import re
from copy import copy
from datetime import datetime as dt


# -------------------------------------------------
# DATA LOADING
# -------------------------------------------------

def clean_columns(df):
    df.columns = df.columns.str.strip()
    return df


def load_main_data(first_file_path, second_file_path):
    df1 = clean_columns(pd.read_excel(first_file_path, sheet_name=0))
    df2_all = pd.read_excel(second_file_path, sheet_name=None)
    for key in df2_all:
        df2_all[key] = clean_columns(df2_all[key])
    return df1, df2_all


def load_support_sheets(first_file_path):
    mont_df = clean_columns(pd.read_excel(first_file_path, sheet_name="MONT"))
    trykktest_df = clean_columns(pd.read_excel(first_file_path, sheet_name="Trykktest"))
    prikling_df = clean_columns(pd.read_excel(first_file_path, sheet_name="Prikling"))
    return mont_df, trykktest_df, prikling_df


# -------------------------------------------------
# LOOKUPS
# -------------------------------------------------

def get_trykktest_prodno(size, length, trykktest_df):
    if size is None:
        return None

    if length < 3000:
        mapping = {
            ("04", "06", "08"): 90094,
            ("10", "12", "16"): 90095,
            ("20", "24"): 90096,
            ("32",): 90097
        }
    else:
        mapping = {
            ("04", "06", "08"): 90098,
            ("10", "12", "16"): 90099,
            ("20", "24"): 900101,
            ("32",): 900102
        }

    prod_no = None
    for keys, val in mapping.items():
        if size in keys:
            prod_no = val
            break

    if prod_no is None:
        return None

    row = trykktest_df.loc[trykktest_df["Prod.no"] == prod_no]
    return row.iloc[0] if not row.empty else None


def get_prikling_row(size, prikling_df):
    if size is None:
        return None
    
    if size in ["04", "06", "08", "10"]:
        prod_no = 90015
    elif size in ["12", "16"]:
        prod_no = 90016
    elif size in ["20", "24", "32"]:
        prod_no = 90017
    else:
        return None
    
    row = prikling_df.loc[prikling_df["Prod.no"] == prod_no]
    return row.iloc[0] if not row.empty else None


def get_mont_row(size, sheet_key, mont_df):
    if size is None:
        return None
    
    size = str(size).strip()
    sheet_key = str(sheet_key).lower()

    if len(mont_df) < 4:
        return None

    if size in ["04", "06", "08", "10"]:
        return mont_df.iloc[0]

    if "316" in sheet_key and "5" not in sheet_key:
        if size in ["12", "16"]:
            return mont_df.iloc[1]
        if size in ["20", "24", "32"]:
            return mont_df.iloc[2]

    if "5-316" in sheet_key:
        return mont_df.iloc[3]

    if "st" in sheet_key:
        if size in ["12", "16"]:
            return mont_df.iloc[1]
        if size in ["20", "24", "32"]:
            return mont_df.iloc[2]

    return None


def _extract_sheet_key_from_sheetname(sheet_name):
    """Extract sheet key from sheet name"""
    if not sheet_name:
        return ""
    m = re.search(r"\(([^)]+)\)", sheet_name)
    if m:
        return f"({m.group(1)})"
    if "316" in sheet_name:
        return "(316)"
    if "GSM" in sheet_name or "GS" in sheet_name:
        return "(GSM)" if "GSM" in sheet_name else "(GS)"
    return ""


def _multiply_row_quantity(row, multiplier):
    """Multiply the quantity (4th column) of a row"""
    if len(row) < 4:
        return row
    val = row[3]
    if val == "" or val is None:
        return row
    try:
        num = float(val)
    except Exception:
        return row
    num *= multiplier
    if abs(num - round(num)) < 1e-9:
        num = int(round(num))
    else:
        num = round(num, 3)
    row[3] = num
    return row


# -------------------------------------------------
# SUMMARY PARSING
# -------------------------------------------------

def find_matches_from_summary(first_line, df1, df2_all, material_pref=None):
    """Parse summary line and find matching rows from dataframes"""
    part1 = part2 = part3 = part4 = angle = None
    length_int = None
    
    s = first_line.strip()
    s = s.replace("°", "")
    parts = s.split("/")
    
    if len(parts) >= 4:
        part1, part2, part3, part4 = parts[0], parts[1], parts[2], parts[3]
        if len(parts) >= 5:
            angle = parts[4]
    else:
        if len(parts) >= 2:
            part1 = parts[0]
            part2 = parts[1]
        if len(parts) >= 3:
            part3 = parts[2]
        if len(parts) >= 4:
            part4 = parts[3]

    try:
        length_int = int(re.sub(r'\D', '', part2)) if part2 is not None else None
    except Exception:
        length_int = None

    # Find selected_first_row
    selected_row = None
    if part1:
        for _, row in df1.iterrows():
            b = str(row.get("Beskrivelse", "")).strip()
            b2 = str(row.get("Beskrivelse_2", "")).strip()
            if b.startswith(part1) or b2.startswith(part1) or part1 in b2 or part1 in b:
                selected_row = row
                break

    second_row1 = second_row2 = None
    sheet_name_found = None
    size_str = None

    def norm_key(x):
        return str(x).strip()

    preferred_marker = None
    if material_pref:
        mp = material_pref.lower()
        if "syre" in mp or "316" in mp:
            preferred_marker = "316"
        elif "stål" in mp or "stal" in mp or "st" in mp:
            preferred_marker = "st"

    candidate_sheets = []
    for sheet_name, df in df2_all.items():
        dfc = clean_columns(df) if isinstance(df, pd.DataFrame) else df
        found1 = None
        found2 = None
        for _, r in dfc.iterrows():
            desc = norm_key(r.get("Beskrivelse", ""))
            if part3 and (desc.startswith(part3) or part3 in desc):
                found1 = r
            if part4 and (desc.startswith(part4) or part4 in desc):
                found2 = r
            if found1 is not None and found2 is not None:
                break
        if found1 is not None and found2 is not None:
            candidate_sheets.append((sheet_name, found1, found2))

    if candidate_sheets:
        picked = None
        if preferred_marker:
            for sheet_name, f1, f2 in candidate_sheets:
                if preferred_marker in sheet_name:
                    picked = (sheet_name, f1, f2)
                    break
        if not picked:
            picked = candidate_sheets[0]
        sheet_name_found, second_row1, second_row2 = picked
        m = re.match(r"Kuplinger\s+(\d{1,3})", sheet_name_found)
        if m:
            size_str = m.group(1)
            if len(size_str) < 2:
                size_str = size_str.zfill(2)
        return selected_row, second_row1, second_row2, sheet_name_found, size_str, length_int

    for sheet_name, df in df2_all.items():
        dfc = clean_columns(df) if isinstance(df, pd.DataFrame) else df
        if second_row1 is None and part3:
            for _, r in dfc.iterrows():
                desc = norm_key(r.get("Beskrivelse", ""))
                if desc.startswith(part3) or part3 in desc:
                    second_row1 = r
                    sheet_name_found = sheet_name if sheet_name_found is None else sheet_name_found
                    if size_str is None:
                        m = re.match(r"Kuplinger\s+(\d{1,3})", sheet_name)
                        if m:
                            size_str = m.group(1)
                            if len(size_str) < 2:
                                size_str = size_str.zfill(2)
                    break
        if second_row2 is None and part4:
            for _, r in dfc.iterrows():
                desc = norm_key(r.get("Beskrivelse", ""))
                if desc.startswith(part4) or part4 in desc:
                    second_row2 = r
                    if sheet_name_found is None:
                        sheet_name_found = sheet_name
                        if size_str is None:
                            m = re.match(r"Kuplinger\s+(\d{1,3})", sheet_name)
                            if m:
                                size_str = m.group(1)
                                if len(size_str) < 2:
                                    size_str = size_str.zfill(2)
                    break

    return selected_row, second_row1, second_row2, sheet_name_found, size_str, length_int


# -------------------------------------------------
# CERTIFICATE DATA
# -------------------------------------------------

def fill_pressure_test_certificate_data(
    pressure_details,
    selected_row,
    second_rows,
    size_str,
    length_int,
    material
):
    """Fill pressure test certificate with data"""
    current_date = dt.now().strftime("%d.%m.%Y")

    try:
        working_pressure_val = float(selected_row.get("Trykk(bar)", 0)) if selected_row is not None else 0
    except:
        working_pressure_val = 0

    burst_pressure_val = working_pressure_val * 4
    test_pressure_val = working_pressure_val * 1.5

    part1 = str(selected_row["Beskrivelse"])[:7] if selected_row is not None else ""
    part2 = str(length_int if length_int else "")
    part3 = str(second_rows[0]["Beskrivelse"])[:9 if material == "stål" else 15] if second_rows[0] is not None else ""
    part4 = str(second_rows[1]["Beskrivelse"])[:9 if material == "stål" else 15] if second_rows[1] is not None else ""

    angle = pressure_details.get("angle", "")
    
    if angle:
        hose_spec = f"{part1}/{part2}/{part3}/{part4}/{angle}°"
    else:
        hose_spec = f"{part1}/{part2}/{part3}/{part4}"

    couplings_str = ""
    if second_rows[0] is not None and second_rows[1] is not None:
        coup1 = str(second_rows[0]["Beskrivelse"])[:9 if material == "stål" else 15]
        coup2 = str(second_rows[1]["Beskrivelse"])[:9 if material == "stål" else 15]
        couplings_str = f"{coup1} / {coup2}"
    elif second_rows[0] is not None:
        coup1 = str(second_rows[0]["Beskrivelse"])[:9 if material == "stål" else 15]
        couplings_str = coup1
    elif second_rows[1] is not None:
        coup2 = str(second_rows[1]["Beskrivelse"])[:9 if material == "stål" else 15]
        couplings_str = coup2

    certificate_data = {
        "A7": pressure_details.get("kunde", ""),
        "A10": pressure_details.get("kundens_best_nr", ""),
        "E10": pressure_details.get("hydra_ordre_nr", ""),
        "A13": pressure_details.get("kundes_del_nr", ""),
        "A16": hose_spec,
        "A19": str(size_str),
        "A22": str(length_int if length_int else ""),
        "A25": couplings_str,
        "B28": current_date,
        "B31": current_date,
        "A34": f"{working_pressure_val:.1f}",
        "D34": f"{burst_pressure_val:.1f}",
        "G34": f"{test_pressure_val:.1f}",
        "A40": str(pressure_details.get("antall_slanger", 1)),
    }

    return certificate_data


# -------------------------------------------------
# EXCEL OUTPUT
# -------------------------------------------------

def copy_sheet_with_formatting(source_wb, source_sheet_name, target_wb, target_sheet_name):
    """Copy entire sheet with all formatting, images, and structure preserved"""
    source_ws = source_wb[source_sheet_name]
    target_ws = target_wb.create_sheet(target_sheet_name)

    # Copy all cell data and formatting
    for row in source_ws.iter_rows():
        for source_cell in row:
            target_cell = target_ws[source_cell.coordinate]
            target_cell.value = source_cell.value

            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.border = copy(source_cell.border)
                target_cell.fill = copy(source_cell.fill)
                target_cell.number_format = copy(source_cell.number_format)
                target_cell.protection = copy(source_cell.protection)
                target_cell.alignment = copy(source_cell.alignment)

    # Copy merged cells
    for merged_range in source_ws.merged_cells.ranges:
        target_ws.merge_cells(str(merged_range))

    # Copy column widths
    for col_letter, col_dimension in source_ws.column_dimensions.items():
        target_col = target_ws.column_dimensions[col_letter]
        target_col.width = col_dimension.width

    # Copy row heights
    for row_num, row_dimension in source_ws.row_dimensions.items():
        target_row = target_ws.row_dimensions[row_num]
        target_row.height = row_dimension.height

    # Copy images/drawings
    for image in source_ws._images:
        target_ws.add_image(copy(image), image.anchor)

    # Copy page setup and print settings
    target_ws.page_setup = source_ws.page_setup
    target_ws.page_margins = copy(source_ws.page_margins)

    return target_ws


def create_output_workbook(output_rows):
    """Create output workbook with data"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Output"

    headers = ["Prod.no", "Beskrivelse", "Lager", "Antall"]

    for col_num, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_num).value = header

    for row_num, row_data in enumerate(output_rows, 2):
        for col_num, val in enumerate(row_data, 1):
            ws.cell(row=row_num, column=col_num).value = val

    # Set column widths
    for col_num in range(1, 5):
        col_letter = openpyxl.utils.get_column_letter(col_num)
        ws.column_dimensions[col_letter].width = 20

    return wb


def add_certificate_sheet(output_wb, template_path, certificate_data, sheet_name):
    """Add certificate sheet from template"""
    template_wb = openpyxl.load_workbook(template_path)
    
    cert_ws = copy_sheet_with_formatting(
        template_wb,
        template_wb.sheetnames[0],
        output_wb,
        sheet_name
    )

    # Fill in certificate data
    for cell_ref, value in certificate_data.items():
        try:
            cert_ws[cell_ref].value = value
        except Exception:
            pass

    return output_wb


def add_sluttkontroll_sheet(output_wb, template_path, kunde="", hydra_ordre_nr=""):
    """Add Sluttkontroll sheet from template"""
    template_wb = openpyxl.load_workbook(template_path)
    
    slutt_ws = copy_sheet_with_formatting(
        template_wb,
        template_wb.sheetnames[0],
        output_wb,
        "Sluttkontroll Slanger"
    )

    # Fill in data
    try:
        slutt_ws["B7"] = kunde
        slutt_ws["B9"] = hydra_ordre_nr
    except Exception:
        pass

    return output_wb