# data_loader.py
import pandas as pd
import streamlit as st
import json
import os
from datetime import date
from constants import FILE_PATH, MONTH_MAP

def _empty_dfs():
    return (pd.DataFrame(columns=["EmployeeName","JobTitle","Department","Manager"]),
            pd.DataFrame(columns=["JobTitle","KPI_Name","Weight"]),
            pd.DataFrame(columns=["EmployeeName","Month","KPI_Name","Weight","KPI_%","Evaluator","Notes","Year","EvalDate","Training"]))

def migrate_from_excel():
    if not os.path.exists(FILE_PATH):
        return False, f"ملف {FILE_PATH} غير موجود"
    try:
        df_emp = pd.read_excel(FILE_PATH, sheet_name="EMPLOYEES")
        df_emp.columns = [str(c).strip() for c in df_emp.columns]
        employees_list = []
        for _, row in df_emp.iterrows():
            try:
                dept_val = str(row.get("القسم", row.get("Department", "")))
                mgr_val = str(row.get("اسم المقيم", row.get("Manager", "")))
                employees_list.append({
                    "EmployeeName": str(row.get("EmployeeName", "")),
                    "JobTitle": str(row.get("JobTitle", "")),
                    "Department": dept_val,
                    "Manager": mgr_val
                })
            except: pass
        with open("emp_profiles.json","w",encoding="utf-8") as f:
            json.dump({"employees": employees_list}, f, ensure_ascii=False, indent=2)

        df_kpi = pd.read_excel(FILE_PATH, sheet_name="KPIs")
        df_kpi.columns = [str(c).strip() for c in df_kpi.columns]
        kpis_list = []
        for _, row in df_kpi.iterrows():
            try:
                kpis_list.append({
                    "JobTitle": str(row.get("JobTitle", "")),
                    "KPI_Name": str(row.get("KPI_Name", "")),
                    "Weight": float(row.get("Weight", 0))
                })
            except: pass
        os.makedirs("db", exist_ok=True)
        with open("db/kpis.json","w",encoding="utf-8") as f:
            json.dump({"kpis": kpis_list}, f, ensure_ascii=False, indent=2)
        
        load_data.clear()
        return True, f"تم نقل {len(employees_list)} موظف و {len(kpis_list)} مؤشر"
    except Exception as e:
        return False, f"فشل النقل: {e}"

@st.cache_data(ttl=30)
def load_data():
    try:
        if os.path.exists("emp_profiles.json") and os.path.exists("db/kpis.json"):
            with open("emp_profiles.json","r",encoding="utf-8") as f:
                emp_json = json.load(f)
                df_emp = pd.DataFrame(emp_json.get("employees", []))
            with open("db/kpis.json","r",encoding="utf-8") as f:
                kpi_json = json.load(f)
                df_kpi = pd.DataFrame(kpi_json.get("kpis", []))
            
            ev_path = "db/evaluations.json"
            if not os.path.exists(ev_path):
                os.makedirs("db", exist_ok=True)
                with open(ev_path, "w", encoding="utf-8") as f:
                    json.dump({"evaluations": []}, f)
            
            with open(ev_path, "r", encoding="utf-8") as f:
                ev_json = json.load(f)
            df_data = pd.DataFrame(ev_json.get("evaluations", []))
            
            if not df_emp.empty and not df_kpi.empty:
                return df_emp, df_kpi, df_data
    except Exception as e:
        st.warning(f"خطأ في قراءة JSON: {e}")

    if os.path.exists(FILE_PATH):
        try:
            df_emp = pd.read_excel(FILE_PATH, sheet_name="EMPLOYEES")
            df_emp.columns = [str(c).strip() for c in df_emp.columns]
            df_kpi = pd.read_excel(FILE_PATH, sheet_name="KPIs")
            df_kpi.columns = [str(c).strip() for c in df_kpi.columns]
            return df_emp, df_kpi, _empty_dfs()[2]
        except:
            pass
    
    return _empty_dfs()


# ============================================================
# دالة الحفظ - مبرمجة من الصفر
# ============================================================
def save_evaluation(emp_name, month_ar, year, manager, dept, kpi_rows, notes="", training=""):
    """
    حفظ تقييم موظف في ملف JSON
    """
    # 1. تجهيز المسار
    folder = "db"
    file_path = os.path.join(folder, "evaluations.json")
    
    # 2. إنشاء المجلد إذا لم يكن موجوداً
    if not os.path.exists(folder):
        os.makedirs(folder)
    
    # 3. قراءة التقييمات الحالية أو إنشاء قائمة فارغة
    all_evaluations = []
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()
                if content.strip():
                    data = json.loads(content)
                    all_evaluations = data.get("evaluations", [])
        except:
            all_evaluations = []
    
    # 4. تحويل اسم الشهر
    month_en = MONTH_MAP.get(month_ar, month_ar)
    
    # 5. تاريخ اليوم
    today_str = date.today().strftime("%Y-%m-%d")
    
    # 6. إضافة صفوف التقييم
    count = 0
    for item in kpi_rows:
        kpi_name = str(item[0])
        weight = float(item[1])
        grade = float(item[2])
        
        all_evaluations.append({
            "EmployeeName": str(emp_name),
            "Month": str(month_en),
            "Year": int(year),
            "KPI_Name": kpi_name,
            "Weight": weight,
            "KPI_%": grade,
            "Evaluator": str(manager),
            "Notes": str(notes) if notes else "",
            "Training": str(training) if training else "",
            "EvalDate": today_str
        })
        count += 1
    
    # 7. حفظ الملف
    try:
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump({"evaluations": all_evaluations}, f, ensure_ascii=False, indent=2)
        
        # 8. مسح الذاكرة المؤقتة
        load_data.clear()
        
        return True, f"تم حفظ {count} صف تقييم بنجاح"
    except Exception as e:
        return False, str(e)


def get_emp_notes(emp):
    return "", ""
