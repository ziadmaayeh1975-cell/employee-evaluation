"""
disciplinary_loader.py — تحميل الإجراءات التأديبية من ملف Excel
"""
import pandas as pd
import streamlit as st
from datetime import datetime

DISCIPLINARY_FILE = "الاجراءات التاديبية.xls"

@st.cache_data(ttl=300)
def load_disciplinary_actions():
    """
    تحميل الإجراءات التأديبية من ملف Excel
    """
    try:
        df = pd.read_excel(DISCIPLINARY_FILE, engine="openpyxl")
        
        # تنظيف أسماء الأعمدة
        df.columns = [str(c).strip() for c in df.columns]
        
        # توحيد اسماء الأعمدة المتوقعة
        expected_columns = {
            "عدد ايام الخصم": "deduction_days",
            "نوع الإنذار": "warning_type", 
            "تاريخ الإنذار": "warning_date",
            "سبب الإنذار": "reason",
            "اسم الموظف": "employee_name",
            "رقم الموظف": "employee_id"
        }
        
        # إعادة تسمية الأعمدة إذا وجدت
        for old, new in expected_columns.items():
            if old in df.columns:
                df.rename(columns={old: new}, inplace=True)
        
        # تحويل تاريخ الإنذار إلى تاريخ فقط
        if "warning_date" in df.columns:
            df["warning_date"] = pd.to_datetime(df["warning_date"]).dt.date
        
        # تنظيف الأسماء
        if "employee_name" in df.columns:
            df["employee_name"] = df["employee_name"].astype(str).str.strip()
        
        return df
    
    except Exception as e:
        st.error(f"❌ خطأ في تحميل ملف الإجراءات التأديبية: {e}")
        return pd.DataFrame(columns=["employee_name", "warning_type", "warning_date", "reason", "deduction_days"])


def get_employee_disciplinary(df_disc, emp_name, year=None, month=None):
    """
    جلب الإجراءات التأديبية لموظف معين
    يمكن تصفيتها حسب السنة والشهر
    """
    if df_disc.empty or "employee_name" not in df_disc.columns:
        return pd.DataFrame()
    
    # فلترة حسب اسم الموظف
    emp_actions = df_disc[df_disc["employee_name"].str.contains(emp_name, na=False, case=False)]
    
    if emp_actions.empty:
        return pd.DataFrame()
    
    # فلترة حسب السنة
    if year and "warning_date" in emp_actions.columns:
        emp_actions = emp_actions[pd.to_datetime(emp_actions["warning_date"]).dt.year == int(year)]
    
    # فلترة حسب الشهر
    if month and year and "warning_date" in emp_actions.columns:
        emp_actions = emp_actions[pd.to_datetime(emp_actions["warning_date"]).dt.month == int(month)]
    
    # ترتيب حسب التاريخ (الأحدث أولاً)
    if "warning_date" in emp_actions.columns:
        emp_actions = emp_actions.sort_values("warning_date", ascending=False)
    
    return emp_actions


def format_disciplinary_text(actions_df):
    """
    تنسيق الإجراءات التأديبية كنص لعرضه في التقرير
    """
    if actions_df.empty:
        return "لا يوجد إجراءات تأديبية"
    
    lines = []
    for _, row in actions_df.iterrows():
        warning_date = row.get("warning_date", "")
        warning_type = row.get("warning_type", "")
        reason = row.get("reason", "")
        deduction = row.get("deduction_days", 0)
        
        line = f"• {warning_date}: {warning_type}"
        if deduction > 0:
            line += f" (خصم {deduction} أيام)"
        if reason:
            line += f" - {reason}"
        lines.append(line)
    
    return "\n".join(lines)
