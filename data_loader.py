"""
data_loader.py — يقرأ من قاعدة البيانات أولاً، Excel احتياطياً
مع دعم الأعمدة ديناميكياً في شيتي EMPLOYEES و KPIs
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


def _map_employee_columns(df_emp):
    """إعادة تسمية أعمدة EMPLOYEES ديناميكياً"""
    mapping = {}
    for col in df_emp.columns:
        col_lower = str(col).strip().lower()
        if "رقم الموظف" in col_lower or "employeeid" in col_lower or "emp_id" in col_lower:
            mapping[col] = "رقم الموظف"
        elif "employeename" in col_lower or "اسم الموظف" in col_lower:
            mapping[col] = "EmployeeName"
        elif "jobtitle" in col_lower or "الوظيفة" in col_lower:
            mapping[col] = "JobTitle"
        elif "قسم" in col_lower or "department" in col_lower:
            mapping[col] = "القسم"
        elif "مقيم" in col_lower or "evaluator" in col_lower:
            mapping[col] = "اسم المقيم"
    return df_emp.rename(columns=mapping)


def _map_kpi_columns(df_kpi):
    """إعادة تسمية أعمدة KPIs ديناميكياً بغض النظر عن الترتيب"""
    mapping = {}
    for col in df_kpi.columns:
        col_lower = str(col).strip().lower()
        if "jobtitle" in col_lower or "الوظيفة" in col_lower:
            mapping[col] = "JobTitle"
        elif "kpi_name" in col_lower or "اسم المؤشر" in col_lower or "المؤشر" in col_lower:
            mapping[col] = "KPI_Name"
        elif "weight" in col_lower or "الوزن" in col_lower:
            mapping[col] = "Weight"
    return df_kpi.rename(columns=mapping)


@st.cache_data(ttl=30)
def load_data():
    if _DB_AVAILABLE and db_exists():
        return load_data_from_db()
    
    try:
        import openpyxl
        df_emp = pd.read_excel(FILE_PATH, sheet_name="EMPLOYEES")
        df_kpi = pd.read_excel(FILE_PATH, sheet_name="KPIs")
        df_data = pd.read_excel(FILE_PATH, sheet_name="DATA")
        
        for df in [df_emp, df_kpi, df_data]:
            if not df.empty:
                df.columns = [str(c).strip() for c in df.columns]
                for col in df.select_dtypes("object").columns:
                    df[col] = df[col].astype(str).str.strip()
        
        # ✅ إعادة تسمية أعمدة EMPLOYEES
        df_emp = _map_employee_columns(df_emp)
        
        # ✅ إعادة تسمية أعمدة KPIs
        df_kpi = _map_kpi_columns(df_kpi)
        
        # ✅ التأكد من وجود الأعمدة الأساسية
        required_emp_cols = ["رقم الموظف", "EmployeeName", "JobTitle", "القسم", "اسم المقيم"]
        missing_emp = [col for col in required_emp_cols if col not in df_emp.columns]
        if missing_emp:
            st.error(f"⚠️ الأعمدة المطلوبة غير موجودة في EMPLOYEES: {missing_emp}")
            return None, None, None
        
        required_kpi_cols = ["JobTitle", "KPI_Name", "Weight"]
        missing_kpi = [col for col in required_kpi_cols if col not in df_kpi.columns]
        if missing_kpi:
            st.error(f"⚠️ الأعمدة المطلوبة غير موجودة في KPIs: {missing_kpi}")
            return None, None, None
        
        # ترتيب الأعمدة
        df_emp = df_emp[required_emp_cols]
        df_kpi = df_kpi[required_kpi_cols]
        
        # معالجة DATA
        if "Nots" in df_data.columns and "Notes" not in df_data.columns:
            df_data.rename(columns={"Nots": "Notes"}, inplace=True)
        
        if "Year" not in df_data.columns:
            df_data["Year"] = 2025
        else:
            df_data["Year"] = pd.to_numeric(df_data["Year"], errors="coerce").fillna(2025).astype(int)
        
        for col in ["EvalDate", "Training", "Notes"]:
            if col not in df_data.columns:
                df_data[col] = ""
        
        return df_emp, df_kpi, df_data
    
    except Exception as e:
        st.error(f"❌ تعذّر تحميل البيانات: {e}")
        return None, None, None


def save_evaluation(emp_name, month_ar, year, manager, dept,
                    kpi_rows, notes="", training=""):
    if _DB_AVAILABLE and db_exists():
        ok, err = save_evaluation_to_db(emp_name, month_ar, year, manager, dept,
                                         kpi_rows, notes, training)
        if ok:
            load_data.clear()
        return ok, err
    
    try:
        import openpyxl
        wb = openpyxl.load_workbook(FILE_PATH, keep_vba=True)
        ws = wb["DATA"]
        
        if ws.cell(1, 8).value != "Year":
            ws.cell(1, 8).value = "Year"
        if ws.cell(1, 9).value != "EvalDate":
            ws.cell(1, 9).value = "EvalDate"
        if ws.cell(1, 10).value != "Training":
            ws.cell(1, 10).value = "Training"
        
        month_en = MONTH_MAP.get(month_ar, month_ar)
        eval_date = date.today().strftime("%d/%m/%Y")
        nr = ws.max_row + 1
        
        for item in kpi_rows:
            if len(item) == 4:
                kpi_name, weight, grade, rating_lbl = item
            else:
                kpi_name, weight, grade = item[:3]
                rating_lbl = ""
            
            ws.cell(nr, 1).value = emp_name
            ws.cell(nr, 2).value = month_en
            ws.cell(nr, 3).value = kpi_name
            ws.cell(nr, 4).value = float(weight)
            ws.cell(nr, 5).value = round(float(grade), 2)
            ws.cell(nr, 6).value = manager
            ws.cell(nr, 7).value = notes
            ws.cell(nr, 8).value = int(year)
            ws.cell(nr, 9).value = eval_date
            ws.cell(nr, 10).value = training
            ws.cell(nr, 11).value = rating_lbl
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
        notes = ws["B20"].value or ""
        training = ws["B21"].value or ""
        wb.close()
        return str(notes), str(training)
    except:
        return "", ""
