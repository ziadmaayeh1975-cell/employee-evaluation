"""
app.py — النسخة النهائية المصححة
"""
import streamlit as st
from auth import ensure_session_state, load_users, render_login, stop_if_trial_expired
from data_loader import load_data
from navigation import render_sidebar, render_page_header
from styles import apply_global_styles

# استيراد الدوال مباشرة من الملفات (بدون pages)
from entry import render_entry
from manage import render_manage
from employee_report import render_employee_report
from department_report import render_department_report
from yearly_summary import render_yearly_summary
from settings_page import render_settings

st.set_page_config(page_title="نظام تقييم فنون", layout="wide", page_icon="📊")
apply_global_styles()
ensure_session_state()

users = load_users()
if not st.session_state.logged_in:
    render_login(users)
    st.stop()

stop_if_trial_expired()

# ── مزامنة تلقائية من Excel إذا تم تعديله ──────────────
try:
    from database_manager import sync_from_excel_if_updated, db_exists
    if db_exists():
        synced, sync_msg = sync_from_excel_if_updated()
        if synced and sync_msg:
            st.sidebar.success(sync_msg)
except Exception:
    pass

df_emp, df_kpi, df_data = load_data()
if df_emp is None:
    st.stop()

page = render_sidebar(users)
render_page_header(page)

if page == "entry":
    render_entry(df_emp, df_kpi, df_data)
elif page == "manage":
    render_manage(df_emp, df_kpi, df_data)
elif page == "rep_emp":
    render_employee_report(df_emp, df_kpi, df_data)
elif page == "rep_dept":
    render_department_report(df_emp, df_kpi, df_data)
elif page == "rep_year":
    render_yearly_summary(df_emp, df_kpi, df_data)
elif page == "settings":
    render_settings(df_emp, df_kpi, df_data)
