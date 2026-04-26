"""
data_loader.py — يقرأ من Supabase أولاً، ثم JSON، ثم Excel، وإذا لم يجد شيئاً يُعيد بيانات فارغة
"""
import pandas as pd
import streamlit as st
from datetime import date
from constants import FILE_PATH, MONTH_MAP

try:
    from database_manager import load_data_from_db, save_evaluation_to_db, db_exists
    _DB_AVAILABLE = True
except ImportError:
    _DB_AVAILABLE = False


def _empty_dfs():
    """يُعيد DataFrames فارغة عند عدم وجود بيانات — يسمح للبرنامج بالعمل."""
    df_emp  = pd.DataFrame(columns=["EmployeeName","JobTitle","القسم","اسم المقيم "])
    df_kpi  = pd.DataFrame(columns=["JobTitle","KPI_Name","Weight"])
    df_data = pd.DataFrame(columns=["EmployeeName","Month","KPI_Name","Weight",
                                     "KPI_%","Evaluator","Notes","Year","EvalDate","Training"])
    return df_emp, df_kpi, df_data


@st.cache_data(ttl=30)
def load_data():
    # ① Supabase / JSON
    if _DB_AVAILABLE and db_exists():
        result = load_data_from_db()
        if result[0] is not None:
            return result

    # ② Excel احتياطي
    try:
        import openpyxl, os
        if not os.path.exists(FILE_PATH):
            raise FileNotFoundError(f"{FILE_PATH} غير موجود")

        df_emp  = pd.read_excel(FILE_PATH, sheet_name="EMPLOYEES")
        df_kpi  = pd.read_excel(FILE_PATH, sheet_name="KPIs")
        df_data = pd.read_excel(FILE_PATH, sheet_name="DATA")

        for df in [df_emp, df_kpi, df_data]:
            df.columns = [str(c).strip() for c in df.columns]
            for col in df.select_dtypes("object").columns:
                df[col] = df[col].astype(str).str.strip()

        col_map_emp = {}
        for c in df_emp.columns:
            if c in ("القسم","Department"):   col_map_emp[c] = "Department"
            elif c in ("اسم المقيم","Manager","Evaluator"): col_map_emp[c] = "Manager"
        if col_map_emp:
            df_emp.rename(columns=col_map_emp, inplace=True)

        if "Nots" in df_data.columns and "Notes" not in df_data.columns:
            df_data.rename(columns={"Nots":"Notes"}, inplace=True)

        if "Year" not in df_data.columns:
            df_data["Year"] = 2025
        else:
            df_data["Year"] = pd.to_numeric(df_data["Year"],errors="coerce").fillna(2025).astype(int)
        for col in ["EvalDate","Training","Notes"]:
            if col not in df_data.columns: df_data[col] = ""

        return df_emp, df_kpi, df_data

    except FileNotFoundError:
        # ③ لا يوجد Excel ولا قاعدة بيانات — ابدأ فارغاً
        return _empty_dfs()

    except Exception as e:
        st.warning(f"⚠️ تعذّر تحميل البيانات: {e} — يمكنك استيراد البيانات من الإعدادات.")
        return _empty_dfs()


def save_evaluation(emp_name, month_ar, year, manager, dept,
                    kpi_rows, notes="", training=""):
    if _DB_AVAILABLE and db_exists():
        ok, err = save_evaluation_to_db(emp_name, month_ar, year, manager, dept,
                                         kpi_rows, notes, training)
        if ok: load_data.clear()
        return ok, err
    try:
        import openpyxl
        wb  = openpyxl.load_workbook(FILE_PATH, keep_vba=True)
        ws  = wb["DATA"]
        if ws.cell(1,8).value  != "Year":     ws.cell(1,8).value  = "Year"
        if ws.cell(1,9).value  != "EvalDate": ws.cell(1,9).value  = "EvalDate"
        if ws.cell(1,10).value != "Training": ws.cell(1,10).value = "Training"
        month_en  = MONTH_MAP.get(month_ar, month_ar)
        eval_date = date.today().strftime("%d/%m/%Y")
        nr = ws.max_row + 1
        for item in kpi_rows:
            kpi_name, weight, grade = item[:3]
            rating_lbl = item[3] if len(item) > 3 else ""
            ws.cell(nr,1).value = emp_name;  ws.cell(nr,2).value = month_en
            ws.cell(nr,3).value = kpi_name;  ws.cell(nr,4).value = float(weight)
            ws.cell(nr,5).value = round(float(grade),2)
            ws.cell(nr,6).value = manager;   ws.cell(nr,7).value = notes
            ws.cell(nr,8).value = int(year); ws.cell(nr,9).value = eval_date
            ws.cell(nr,10).value = training; ws.cell(nr,11).value = rating_lbl
            nr += 1
        wb.save(FILE_PATH)
        return True, None
    except Exception as e:
        return False, str(e)


def get_emp_notes(emp_name):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(FILE_PATH, read_only=True, keep_vba=True)
        ws = wb["INPUT"]
        notes    = ws["B20"].value or ""
        training = ws["B21"].value or ""
        wb.close()
        return str(notes), str(training)
    except:
        return "", ""
