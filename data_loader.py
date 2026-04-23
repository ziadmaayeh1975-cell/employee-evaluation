"""
data_loader.py — يقرأ من JSON أولاً، ثم Excel للاستيراد مرة واحدة
"""
import pandas as pd
import streamlit as st
import json
import os
from datetime import date
from constants import FILE_PATH, MONTH_MAP, EMPLOYEES_JSON, KPIS_JSON

def _empty_dfs():
    """يُعيد DataFrames فارغة عند عدم وجود بيانات"""
    df_emp  = pd.DataFrame(columns=["EmployeeName","JobTitle","Department","Manager"])
    df_kpi  = pd.DataFrame(columns=["JobTitle","KPI_Name","Weight"])
    df_data = pd.DataFrame(columns=["EmployeeName","Month","KPI_Name","Weight",
                                     "KPI_%","Evaluator","Notes","Year","EvalDate","Training"])
    return df_emp, df_kpi, df_data


def migrate_from_excel():
    """
    نقل البيانات من Excel إلى JSON - تشغيل مرة واحدة فقط
    """
    if not os.path.exists(FILE_PATH):
        return False, "ملف Excel غير موجود"
    
    try:
        import openpyxl
        
        # 1. قراءة الموظفين
        df_emp = pd.read_excel(FILE_PATH, sheet_name="EMPLOYEES")
        df_emp.columns = [str(c).strip() for c in df_emp.columns]
        
        employees = []
        for _, row in df_emp.iterrows():
            emp = {
                "EmployeeName": str(row.get("EmployeeName", "")),
                "JobTitle": str(row.get("JobTitle", "")),
                "Department": str(row.get("القسم", "")),
                "Manager": str(row.get("اسم المقيم", row.get("اسم المقيم ", "")))
            }
            if emp["EmployeeName"]:
                employees.append(emp)
        
        # حفظ الموظفين
        with open(EMPLOYEES_JSON, "w", encoding="utf-8") as f:
            json.dump({"employees": employees}, f, ensure_ascii=False, indent=4)
        
        # 2. قراءة المؤشرات
        df_kpi = pd.read_excel(FILE_PATH, sheet_name="KPIs")
        df_kpi.columns = [str(c).strip() for c in df_kpi.columns]
        
        kpis_list = []
        for _, row in df_kpi.iterrows():
            kpi = {
                "JobTitle": str(row.get("JobTitle", "")),
                "KPI_Name": str(row.get("KPI_Name", "")),
                "Weight": float(row.get("Weight", 0))
            }
            if kpi["JobTitle"] and kpi["KPI_Name"]:
                kpis_list.append(kpi)
        
        # حفظ المؤشرات
        os.makedirs("db", exist_ok=True)
        with open(KPIS_JSON, "w", encoding="utf-8") as f:
            json.dump({"kpis": kpis_list}, f, ensure_ascii=False, indent=4)
        
        # مسح الكاش
        load_data.clear()
        
        return True, f"✅ تم نقل {len(employees)} موظف و {len(kpis_list)} مؤشر"
        
    except Exception as e:
        return False, f"❌ خطأ: {e}"


def load_from_json():
    """تحميل البيانات من ملفات JSON"""
    df_emp = pd.DataFrame(columns=["EmployeeName","JobTitle","Department","Manager"])
    df_kpi = pd.DataFrame(columns=["JobTitle","KPI_Name","Weight"])
    
    # قراءة الموظفين
    if os.path.exists(EMPLOYEES_JSON):
        try:
            with open(EMPLOYEES_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
                if data.get("employees"):
                    df_emp = pd.DataFrame(data["employees"])
        except:
            pass
    
    # قراءة المؤشرات
    if os.path.exists(KPIS_JSON):
        try:
            with open(KPIS_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
                if data.get("kpis"):
                    df_kpi = pd.DataFrame(data["kpis"])
        except:
            pass
    
    # قراءة التقييمات (DATA)
    data_file = "db/evaluations.json"
    df_data = pd.DataFrame(columns=["EmployeeName","Month","KPI_Name","Weight",
                                     "KPI_%","Evaluator","Notes","Year","EvalDate","Training"])
    
    if os.path.exists(data_file):
        try:
            with open(data_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                if data.get("evaluations"):
                    df_data = pd.DataFrame(data["evaluations"])
        except:
            pass
    else:
        # ملف تقييمات فارغ
        os.makedirs("db", exist_ok=True)
        with open(data_file, "w", encoding="utf-8") as f:
            json.dump({"evaluations": []}, f)
    
    return df_emp, df_kpi, df_data


@st.cache_data(ttl=30)
def load_data():
    """
    تحميل البيانات - يقرأ من JSON أولاً
    إذا كان JSON فارغاً، يحاول القراءة من Excel (لأول مرة فقط)
    """
    # ① محاولة التحميل من JSON
    df_emp, df_kpi, df_data = load_from_json()
    
    # إذا JSON فاضي وجربنا Excel
    if df_emp.empty and df_kpi.empty:
        if os.path.exists(FILE_PATH):
            # نسخة احتياطية من Excel
            try:
                df_emp = pd.read_excel(FILE_PATH, sheet_name="EMPLOYEES")
                df_kpi = pd.read_excel(FILE_PATH, sheet_name="KPIs")
                df_emp.columns = [str(c).strip() for c in df_emp.columns]
                df_kpi.columns = [str(c).strip() for c in df_kpi.columns]
            except:
                return _empty_dfs()
        else:
            return _empty_dfs()
    
    return df_emp, df_kpi, df_data


def save_evaluation(emp_name, month_ar, year, manager, dept,
                    kpi_rows, notes="", training=""):
    """حفظ التقييم في JSON"""
    try:
        data_file = "db/evaluations.json"
        os.makedirs("db", exist_ok=True)
        
        # قراءة الملف الحالي
        evaluations = []
        if os.path.exists(data_file):
            with open(data_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                evaluations = data.get("evaluations", [])
        
        # إضافة التقييمات الجديدة
        month_en = MONTH_MAP.get(month_ar, month_ar)
        eval_date = date.today().strftime("%Y-%m-%d")
        
        for item in kpi_rows:
            kpi_name, weight, grade = item[:3]
            rating_lbl = item[3] if len(item) > 3 else ""
            
            evaluations.append({
                "EmployeeName": emp_name,
                "Month": month_en,
                "KPI_Name": kpi_name,
                "Weight": float(weight),
                "KPI_%": round(float(grade), 2),
                "Evaluator": manager,
                "Notes": notes,
                "Year": int(year),
                "EvalDate": eval_date,
                "Training": training,
                "RatingLabel": rating_lbl
            })
        
        # حفظ الملف
        with open(data_file, "w", encoding="utf-8") as f:
            json.dump({"evaluations": evaluations}, f, ensure_ascii=False, indent=4)
        
        # مسح الكاش
        load_data.clear()
        
        return True, None
        
    except Exception as e:
        return False, str(e)


def get_emp_notes(emp_name):
    """قراءة ملاحظات الموظف"""
    try:
        data_file = "db/evaluations.json"
        if os.path.exists(data_file):
            with open(data_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                for ev in data.get("evaluations", []):
                    if ev.get("EmployeeName") == emp_name:
                        return ev.get("Notes", ""), ev.get("Training", "")
    except:
        pass
    return "", ""
