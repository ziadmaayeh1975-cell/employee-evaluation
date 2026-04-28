import io
from datetime import date
import openpyxl
import pandas as pd
import streamlit as st
from constants import MONTHS_AR, MONTHS_EN, MONTHS_SHORT, MONTH_MAP
from calculations import calc_monthly, get_kpi_avgs, verbal_grade
from data_loader import get_emp_notes
from auth import get_current_reviewer, get_current_role
from report_export import build_employee_sheet, build_summary_sheet, print_preview_html

def _reviewer_emp_set(df_emp):
    """تحديد الموظفين المسموح للمقيم رؤيتهم"""
    role = get_current_role()
    current_reviewer = get_current_reviewer()
    
    # super_admin يرى الكل
    if role == "super_admin":
        return None
    
    # admin بدون reviewer يرى الكل
    if role == "admin" and not current_reviewer:
        return None
    
    # user أو admin مع reviewer يرى موظفيه فقط
    reviewer_col = "اسم المقيم"
    if reviewer_col not in df_emp.columns:
        return None
    
    return set(df_emp[df_emp[reviewer_col].astype(str).str.strip() == current_reviewer]["EmployeeName"].dropna().tolist())

def _reviewer_emp_list(df_emp):
    """قائمة الموظفين المسموح رؤيتهم"""
    allowed = _reviewer_emp_set(df_emp)
    if allowed is None:
        return df_emp["EmployeeName"].dropna().astype(str).str.strip().tolist()
    return list(allowed)

def _safe_df(df):
    """تأكد من وجود الأعمدة المطلوبة في DataFrame"""
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame(columns=["EmployeeName","Month","KPI_Name","Weight","KPI_%","Evaluator","Notes","Year","EvalDate","Training"])
    df = df.copy()
    for col in ["EmployeeName","Month","Year","KPI_%","KPI_Name"]:
        if col not in df.columns:
            df[col] = pd.Series(dtype="object")
    return df

def _get_month_details(df_data, emp_name, month_en, year):
    """استخراج تفاصيل الشهر (تاريخ التقييم، ملاحظات، تدريب)"""
    mask = (df_data["EmployeeName"] == emp_name) & (df_data["Month"] == month_en) & (df_data["Year"] == int(year))
    subset = df_data[mask]
    if subset.empty:
        return "", "", ""
    row = subset.iloc[0]
    
    eval_date = ""
    for col in ["EvalDate","eval_date","EntryDate"]:
        if col in subset.columns and pd.notna(row.get(col, None)):
            val = str(row[col]).strip()
            if val and val not in ("","nan","None"):
                try: 
                    eval_date = pd.to_datetime(val).strftime("%d/%m/%Y")
                except: 
                    eval_date = val
                break
    
    notes = ""
    for col in ["Notes","notes"]:
        if col in subset.columns and pd.notna(row.get(col, None)):
            val = str(row[col]).strip()
            if val and val not in ("","nan","None"):
                notes = val
                break
    
    training = ""
    for col in ["Training","training"]:
        if col in subset.columns and pd.notna(row.get(col, None)):
            val = str(row[col]).strip()
            if val and val not in ("","nan","None"):
                training = val
                break
    
    return eval_date, notes, training

def render_department_report(df_emp, df_kpi, df_data):
    st.subheader("📊 التقرير المفصّل بالأقسام")
    df_data = _safe_df(df_data)
    allowed_emps = _reviewer_emp_set(df_emp)
    
    # تحديد أسماء الأعمدة الصحيحة
    emp_name_col = "EmployeeName"
    dept_col = "القسم"
    
    # فلترة الموظفين حسب الصلاحيات
    if allowed_emps is None:
        dept_emps_all = df_emp
    else:
        dept_emps_all = df_emp[df_emp[emp_name_col].isin(allowed_emps)]
    
    # الحصول على قائمة الأقسام
    dept_list = sorted(dept_emps_all[dept_col].dropna().astype(str).str.strip().unique().tolist())
    
    # واجهة المستخدم
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        sel_dept = st.selectbox("🏢 القسم", ["-- الكل --"] + dept_list, key="dept_sel")
    with col2:
        sel_months = st.multiselect("📅 الأشهر", MONTHS_AR, key="dept_months")
    with col3:
        sel_year = st.selectbox("🗓️ السنة", [2025, 2026, 2027], key="dept_year")
    
    # فلترة الموظفين حسب القسم
    if sel_dept == "-- الكل --":
        dept_emps = dept_emps_all[emp_name_col].dropna().tolist()
    else:
        dept_emps = dept_emps_all[dept_emps_all[dept_col].astype(str).str.strip() == sel_dept][emp_name_col].tolist()
    
    # التحقق من وجود بيانات تقييم
    if df_data.empty or emp_name_col not in df_data.columns:
        st.info("ℹ️ لا توجد تقييمات محفوظة حتى الآن.")
        return
    
    dept_with_data = [e for e in dept_emps if e in df_data[emp_name_col].values]
    if not dept_with_data:
        st.info("لا توجد بيانات تقييم لهذا القسم بعد.")
        return
    
    st.markdown(f"**عدد الموظفين: {len(dept_with_data)}** (لديهم بيانات تقييم)")
    
    # تحويل الأشهر المحددة إلى الإنجليزية للتصفية
    months_en_filter = [MONTH_MAP[m] for m in sel_months] if sel_months else None
    
    # تجميع الملخص
    summary = []
    for emp in dept_with_data:
        # جلب بيانات الموظف من df_emp
        emp_info = df_emp[df_emp[emp_name_col] == emp]
        if emp_info.empty:
            continue
        emp_info = emp_info.iloc[0]
        emp_dept = str(emp_info.get(dept_col, "—"))
        emp_id = str(emp_info.get("رقم الموظف", ""))
        job = str(emp_info.get("JobTitle", ""))
        reviewer = str(emp_info.get("اسم المقيم", ""))
        
        # حساب الدرجات الشهرية
        monthly_scores = []
        for month_en in MONTHS_EN:
            if months_en_filter and month_en not in months_en_filter:
                continue
            score = calc_monthly(df_data, emp, month_en, sel_year)
            monthly_scores.append(score)
        
        active_scores = [s for s in monthly_scores if s > 0]
        avg_score = sum(active_scores) / len(active_scores) if active_scores else 0.0
        avg_percent = round(avg_score * 100, 1)
        
        summary.append({
            "emp": emp,
            "dept": emp_dept,
            "months": len(active_scores),
            "pct": avg_percent,
            "verb": verbal_grade(avg_percent) if active_scores else "—",
            "job": job,
            "reviewer": reviewer,
            "emp_id": emp_id
        })
    
    # ترتيب حسب النسبة تنازلياً
    summary.sort(key=lambda x: x["pct"], reverse=True)
    
    # عرض جدول الملخص
    st.dataframe(pd.DataFrame([{
        "الموظف": s["emp"],
        "القسم": s["dept"],
        "السنة": sel_year,
        "الأشهر": s["months"],
        "المعدل (%)": s["pct"],
        "التقييم": s["verb"],
    } for s in summary]), hide_index=True, use_container_width=True)
    
    # إنشاء ملف Excel للتقرير
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    for s in summary:
        # جلب مؤشرات الأداء للموظف
        kpis = get_kpi_avgs(df_data, df_kpi, s["emp"], s["job"], months_en_filter, sel_year)
        
        # تجهيز البيانات الشهرية
        monthly_rep = []
        for idx, (en, short) in enumerate(zip(MONTHS_EN, MONTHS_SHORT)):
            if months_en_filter and en not in months_en_filter:
                monthly_rep.append((idx+1, short, 0.0, "", "", ""))
            else:
                score = calc_monthly(df_data, s["emp"], en, sel_year)
                eval_date, notes, training = _get_month_details(df_data, s["emp"], en, sel_year)
                monthly_rep.append((idx+1, short, score, eval_date, notes, training))
        
        # جلب الملاحظات والتدريب
        emp_notes = ""
        emp_train = ""
        for item in monthly_rep:
            if item[2] > 0:
                emp_notes = item[4] or ""
                emp_train = item[5] or ""
                break
        if not emp_notes and not emp_train:
            emp_notes, emp_train = get_emp_notes(s["emp"])
        
        # بناء شيت الموظف
        build_employee_sheet(
            wb, s["emp"], s["job"], s["dept"], s["reviewer"], sel_year,
            kpis, monthly_rep, emp_notes, emp_train, employee_id=s["emp_id"]
        )
    
    # إضافة شيت الملخص
    period_label = "، ".join(sel_months) if sel_months else "كل الأشهر"
    sum_title = f"ملخص – {sel_dept if sel_dept != '-- الكل --' else 'الكل'} – {sel_year} – {period_label}"
    build_summary_sheet(
        wb,
        [(s["emp"], s["dept"], s["months"], s["pct"], s["verb"]) for s in summary],
        sum_title,
        year=sel_year
    )
    
    # نقل شيت الملخص إلى البداية
    if len(wb.worksheets) > 1:
        wb.move_sheet(wb.worksheets[-1], offset=-(len(wb.worksheets)-1))
    
    # حفظ وتحميل الملف
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    
    dept_label = sel_dept.replace(" ", "_") if sel_dept != "-- الكل --" else "الكل"
    st.download_button(
        label=f"📥 تحميل Excel ({len(summary)} موظف)",
        data=buf,
        file_name=f"تقارير_{dept_label}_{date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    
    st.markdown("---")
    st.markdown("#### 🖨️ معاينة وطباعة")
    html_preview = print_preview_html(io.BytesIO(buf.getvalue()))
    st.components.v1.html(html_preview, height=1400, scrolling=True)
