import sys
import os
import pandas as pd
from utilities.sheet_header_utils import get_header_index


def load_excel_sheet(filepath, sheet_name, header_row):
    excel_sheet = pd.read_excel(filepath, sheet_name=sheet_name, header=header_row - 1, engine='openpyxl')
    excel_sheet.columns = excel_sheet.columns.astype(str)
    excel_sheet = excel_sheet.loc[:, ~excel_sheet.columns.str.contains(r'^Unnamed', case=False)]
    excel_sheet = excel_sheet.dropna(axis=1, how='all')
    return excel_sheet
def get_row_dict(file_path, sheet_name, header_row, index_value):
    try:
        excel_sheet = load_excel_sheet(file_path, sheet_name, header_row)
        excel_sheet.columns = [clean_column_name(col) for col in excel_sheet.columns]
        match = get_first_match(excel_sheet, index_value)
        return match.to_dict() if match is not None else {}
    except Exception as e:
        print(f"get_row_dict error: {e}")
        return {}

def clean_column_name(name):
    if not isinstance(name, str):
        return name
    name = name.replace('\n', ' ')
    name = name.split('*')[0].strip()
    return name

def get_first_match(excel_sheet, search_term):
    first_col = excel_sheet.columns[0]
    matches = excel_sheet[excel_sheet[first_col].astype(str) == search_term]
    return matches.iloc[0] if not matches.empty else None


def get_saved_header_row(filename, sheet_name, default=2):
    saved = get_header_index(filename, sheet_name)
    return saved if isinstance(saved, int) and saved > 0 else default
