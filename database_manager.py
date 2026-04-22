"""
database_manager.py — قاعدة البيانات المحلية بديل Excel
مع إضافة: الإجراءات التأديبية + معلومات الموظف الإضافية + التقييم السابق
"""
import json, os
import pandas as pd
import streamlit as st
from datetime import date, datetime
from constants import MONTH_MAP, PERSONAL_KPIS, PERSONAL_WEIGHT, MONTHS_AR, MONTHS_EN

DB_DIR  = "db"
EMP_DB  = os.path.join(DB_DIR, "employees.json")
KPI_DB  = os.path.join(DB_DIR, "kpis.json")
DATA_DB = os.path.join(DB_DIR, "evaluations.json")
META_FILE = os.path.join(DB_DIR, "db_meta.json")
DISCIPLINARY_DB = os.path.join(DB_DIR, "disciplinary.json")
EMP_EXTRA_DB = os.path.join(DB_DIR, "employees_extra.json")

os.makedirs(DB_DIR, exist_ok=True)


def db_exists():
    return all(os.path.exists(p) for p in [EMP_DB, KPI_DB, DATA_DB])


def get_db_meta():
    if os.path.exists(META_FILE):
        with open(META_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


@st.cache_data(ttl=30)
def load_data_from_db():
    try:
        with open(EMP_DB,  "r", encoding="utf-8") as f: emp_records  = json.load(f)
        with open(KPI_DB,  "r", encoding="utf-8") as f: kpi_records  = json.load(f)
        with open(DATA_DB, "r", encoding="utf-8") as f: data_records = json.load(f)

        df_emp  = pd.DataFrame(emp_records)
        df_kpi  = pd.DataFrame(kpi_records)
        df_data = pd.DataFrame(data_records)

        for df in [df_emp, df_kpi, df_data]:
            df.columns = [str(c).strip() for c in df.columns]
            for col in df.columns:
                if str(df[col].dtype) in ('object', 'str', 'string'):
                    df[col] = df[col].astype(str).str.strip()

        df_kpi["Weight"] = pd.to_numeric(df_kpi["Weight"], errors="coerce").fillna(0)
        for col in ["KPI_%", "Weight"]:
            if col in df_data.columns:
                df_data[col] = pd.to_numeric(df_data[col], errors="coerce").fillna(0)
        if "Year" in df_data.columns:
            df_data["Year"] = pd.to_numeric(df_data["Year"], errors="coerce").fillna(2025).astype(int)

        if "اسم المقيم " not in df_emp.columns and "اسم المقيم" in df_emp.columns:
            df_emp.rename(columns={"اسم المقيم": "اسم المقيم "}, inplace=True)

        return df_emp, df_kpi, df_data

    except FileNotFoundError:
        return None, None, None
    except Exception as e:
        st.error(f"❌ خطأ في تحميل قاعدة البيانات: {e}")
        return None, None, None


def load_employees_db():
    """تحميل قائمة الموظفين من قاعدة البيانات"""
    if os.path.exists(EMP_DB):
        with open(EMP_DB, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_evaluation_to_db(emp_name, month_ar, year, manager, dept, kpi_rows, notes="", training=""):
    try:
        with open(DATA_DB, "r", encoding="utf-8") as f:
            records = json.load(f)
        month_en = MONTH_MAP.get(month_ar, month_ar)
        eval_date = date.today().strftime("%d/%m/%Y")
        for item in kpi_rows:
            if len(item) == 4:
                kpi_name, weight, grade, rating_lbl = item
            else:
                kpi_name, weight, grade = item[:3]
                rating_lbl = ""
            records.append({
                "EmployeeName": emp_name,
                "Month": month_en,
                "KPI_Name": kpi_name,
                "Weight": float(weight),
                "KPI_%": round(float(grade), 2),
                "RatingLabel": rating_lbl,
                "Evaluator": manager,
                "Nots": notes,
                "Year": int(year),
                "EvalDate": eval_date,
                "Training": training,
            })
        with open(DATA_DB, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
        load_data_from_db.clear()
        return True, None
    except Exception as e:
        return False, str(e)


# ════════════════════════════════════════════════════════════════
# ✅ إجراءات تأديبية
# ════════════════════════════════════════════════════════════════
def load_disciplinary_actions():
    """تحميل جميع الإجراءات التأديبية"""
    if os.path.exists(DISCIPLINARY_DB):
        with open(DISCIPLINARY_DB, "r", encoding="utf-8") as f:
            return 
