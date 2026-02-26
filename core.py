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

    return selected_row, second_row1, second_row2, sheet_name_found, size_str, length_int
