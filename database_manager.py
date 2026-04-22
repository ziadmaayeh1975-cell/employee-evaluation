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
            return json.load(f)
    return []


def save_disciplinary_action(action):
    """حفظ إجراء تأديبي جديد"""
    try:
        actions = load_disciplinary_actions()
        actions.append(action)
        with open(DISCIPLINARY_DB, "w", encoding="utf-8") as f:
            json.dump(actions, f, ensure_ascii=False, indent=2)
        return True, None
    except Exception as e:
        return False, str(e)


def delete_disciplinary_action(action_id):
    """حذف إجراء تأديبي"""
    try:
        actions = load_disciplinary_actions()
        actions = [a for a in actions if a.get("id") != action_id]
        with open(DISCIPLINARY_DB, "w", encoding="utf-8") as f:
            json.dump(actions, f, ensure_ascii=False, indent=2)
        return True, None
    except Exception as e:
        return False, str(e)


def get_employee_disciplinary_actions(emp_name, year=None, month=None):
    """جلب الإجراءات التأديبية لموظف معين"""
    actions = load_disciplinary_actions()
    emp_actions = [a for a in actions if a.get("employee_name") == emp_name]
    
    if year:
        emp_actions = [a for a in emp_actions if str(a.get("year", "")) == str(year)]
    
    return emp_actions


def get_disciplinary_for_month(emp_name, year, month_ar):
    """جلب الإجراءات التأديبية لشهر معين (للتقارير)"""
    actions = load_disciplinary_actions()
    emp_actions = []
    for action in actions:
        if action.get("employee_name") != emp_name:
            continue
        if int(action.get("year", 0)) != int(year):
            continue
        action_month = action.get("month", "")
        if action_month == month_ar or action_month == MONTHS_AR[int(action_month)-1] if action_month.isdigit() else False:
            action_type = action.get("action_type", "")
            description = action.get("description", "")
            emp_actions.append(f"{action_type}: {description}")
    
    return "; ".join(emp_actions) if emp_actions else "-"


def get_previous_evaluation(emp_name, current_year):
    """جلب نتيجة التقييم السابق للموظف"""
    try:
        if not os.path.exists(DATA_DB):
            return None
        
        with open(DATA_DB, "r", encoding="utf-8") as f:
            records = json.load(f)
        
        prev_year = int(current_year) - 1
        prev_records = [r for r in records 
                        if r.get("EmployeeName") == emp_name 
                        and int(r.get("Year", 0)) == prev_year]
        
        if not prev_records:
            return None
        
        total_score = sum(float(r.get("KPI_%", 0)) for r in prev_records) / max(len(prev_records), 1)
        
        if total_score >= 90:
            verbal = "ممتاز"
        elif total_score >= 80:
            verbal = "جيد جداً"
        elif total_score >= 70:
            verbal = "جيد"
        elif total_score >= 60:
            verbal = "مقبول"
        else:
            verbal = "ضعيف"
        
        return {
            "year": prev_year,
            "score": round(total_score, 1),
            "verbal": verbal
        }
    except Exception:
        return None


def import_from_excel(file_path):
    result = {"success": False, "message": "", "counts": {}}
    try:
        df_emp = pd.read_excel(file_path, sheet_name="EMPLOYEES", engine="openpyxl")
        df_kpi = pd.read_excel(file_path, sheet_name="KPIs", engine="openpyxl")
        df_data = pd.read_excel(file_path, sheet_name="DATA", engine="openpyxl")

        for df in [df_emp, df_kpi, df_data]:
            df.columns = [str(c).strip() for c in df.columns]
            for col in df.select_dtypes(include=['object']).columns:
                df[col] = df[col].astype(str).str.strip()

        df_data = df_data.dropna(subset=["EmployeeName","Month","KPI_Name"], how="all")
        df_data = df_data[df_data["EmployeeName"].astype(str).str.strip().ne("nan")]

        if "Year" not in df_data.columns:
            df_data["Year"] = 2025
        else:
            df_data["Year"] = pd.to_numeric(df_data["Year"], errors="coerce").fillna(2025).astype(int)
        
        for col in ["EvalDate","Training","Nots"]:
            if col not in df_data.columns:
                df_data[col] = ""
            else:
                df_data[col] = df_data[col].fillna("").astype(str).replace("nan","")

        df_kpi["Weight"] = pd.to_numeric(df_kpi["Weight"], errors="coerce").fillna(0)
        df_data["Weight"] = pd.to_numeric(df_data["Weight"], errors="coerce").fillna(0)
        df_data["KPI_%"] = pd.to_numeric(df_data["KPI_%"], errors="coerce").fillna(0)

        df_emp = df_emp.dropna(subset=["EmployeeName"])
        df_emp = df_emp[df_emp["EmployeeName"].astype(str).str.strip().ne("nan")]

        emp_records = []
        for _, r in df_emp.iterrows():
            emp_records.append({
                "EmployeeName": str(r.get("EmployeeName","")).strip(),
                "JobTitle": str(r.get("JobTitle","")).strip(),
                "القسم": str(r.get("القسم","")).strip(),
                "اسم المقيم ": str(r.get("اسم المقيم ", r.get("اسم المقيم",""))).strip(),
            })
        with open(EMP_DB, "w", encoding="utf-8") as f:
            json.dump(emp_records, f, ensure_ascii=False, indent=2)

        jobs_total = df_kpi.groupby("JobTitle")["Weight"].sum().to_dict()
        kpi_records = []
        for _, r in df_kpi.iterrows():
            job = str(r["JobTitle"]).strip()
            total = jobs_total.get(job, 100)
            w = float(r["Weight"])
            if total == 100:
                w = round(w * 0.8, 1)
            kpi_records.append({"JobTitle": job, "KPI_Name": str(r["KPI_Name"]).strip(), "Weight": w})
        for job in df_kpi["JobTitle"].unique():
            for p in PERSONAL_KPIS:
                kpi_records.append({"JobTitle": str(job).strip(), "KPI_Name": p, "Weight": float(PERSONAL_WEIGHT)})
        with open(KPI_DB, "w", encoding="utf-8") as f:
            json.dump(kpi_records, f, ensure_ascii=False, indent=2)

        data_records = []
        for _, r in df_data.iterrows():
            data_records.append({
                "EmployeeName": str(r.get("EmployeeName","")).strip(),
                "Month": str(r.get("Month","")).strip(),
                "KPI_Name": str(r.get("KPI_Name","")).strip(),
                "Weight": float(r.get("Weight",0)),
                "KPI_%": float(r.get("KPI_%",0)),
                "Evaluator": str(r.get("Evaluator","")).strip(),
                "Nots": str(r.get("Nots","")).replace("nan","").strip(),
                "Year": int(r.get("Year",2025)),
                "EvalDate": str(r.get("EvalDate","")).replace("nan","").strip(),
                "Training": str(r.get("Training","")).replace("nan","").strip(),
            })
        with open(DATA_DB, "w", encoding="utf-8") as f:
            json.dump(data_records, f, ensure_ascii=False, indent=2)

        meta = {
            "imported_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "source_file": os.path.basename(file_path),
            "employees_count": len(emp_records),
            "kpis_count": len(kpi_records),
            "evaluations_count": len(data_records),
        }
        with open(META_FILE, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)

        load_data_from_db.clear()
        result.update({
            "success": True,
            "message": "✅ تم الاستيراد بنجاح",
            "counts": {
                "employees": len(emp_records),
                "kpis": len(kpi_records),
                "evaluations": len(data_records)
            }
        })
    except Exception as e:
        result["message"] = f"❌ خطأ: {e}"
    return result


def export_db_to_excel():
    import io
    df_emp, df_kpi, df_data = load_data_from_db()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        if df_emp is not None:
            df_emp.to_excel(writer, sheet_name="EMPLOYEES", index=False)
        if df_kpi is not None:
            df_kpi.to_excel(writer, sheet_name="KPIs", index=False)
        if df_data is not None:
            df_data.to_excel(writer, sheet_name="DATA", index=False)
    buf.seek(0)
    return buf.read()


def sync_from_excel_if_updated(file_path="final Apprisal.xlsm"):
    if not os.path.exists(file_path):
        return False, "ملف Excel غير موجود"

    meta = get_db_meta()
    last_import = meta.get("imported_at", "")

    try:
        excel_mtime = os.path.getmtime(file_path)
        excel_dt = datetime.fromtimestamp(excel_mtime)

        if last_import:
            last_dt = datetime.strptime(last_import, "%Y-%m-%d %H:%M")
            if (excel_dt - last_dt).total_seconds() > 60:
                result = import_from_excel(file_path)
                if result["success"]:
                    return True, f"✅ تم تحديث قاعدة البيانات تلقائياً من Excel"
                return False, result["message"]
    except Exception as e:
        return False, str(e)

    return False, ""
