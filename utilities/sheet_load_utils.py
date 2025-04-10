import sys
import os
import pandas as pd
from utilities.sheet_header_utils import get_header_index


def load_excel_sheet(filepath, sheet_name, header_row):
    df = pd.read_excel(filepath, sheet_name=sheet_name, header=header_row - 1, engine='openpyxl')
    df.columns = df.columns.astype(str)
    df = df.loc[:, ~df.columns.str.contains(r'^Unnamed', case=False)]
    df = df.dropna(axis=1, how='all')
    return df

def clean_column_name(name):
    if not isinstance(name, str):
        return name
    name = name.replace('\n', ' ')
    name = name.split('*')[0].strip()
    return name

def get_first_match(df, search_term):
    first_col = df.columns[0]
    matches = df[df[first_col].astype(str) == search_term]
    return matches.iloc[0] if not matches.empty else None


def get_saved_header_row(filename, sheet_name, default=2):
    saved = get_header_index(filename, sheet_name)
    return saved if isinstance(saved, int) and saved > 0 else default
