import io
from datetime import date
import openpyxl
import pandas as pd
import streamlit as st
from constants import MONTHS_EN
from calculations import calc_monthly, calc_yearly, verbal_grade
from auth import get_current_reviewer, get_current_role
from report_export import build_summary_sheet, print_preview_html

def _reviewer_emp_set(df_emp):
    role = get_current_role()
    current_reviewer = get_current_reviewer()
    if role == "super_admin":
        return None
    if role == "admin" and not current_reviewer:
        return None
    reviewer_col = df_emp.columns[3] if len(df_emp.columns) > 3 else df_emp.columns[-1]
    return set(
        str(e).strip() for e in
        df_emp[df_emp[reviewer_col].astype(str).str.strip() == current_reviewer
               ]["EmployeeName"].dropna().tolist()
        if str(e).strip() not in ("","nan")
    )

def _reviewer_emp_list(df_emp):
    allowed = _reviewer_emp_set(df_emp)
    if allowed is None:
        return df_emp["EmployeeName"].dropna().astype(str).str.strip().tolist()
    return list(allowed)

def _safe_df(df):
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame(columns=[
            "
