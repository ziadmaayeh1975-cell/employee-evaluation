# employee_report.py — تقرير الموظف (نسخة بسيطة وآمنة)
import streamlit as st
import pandas as pd
from datetime import date


def render_employee_report(df_emp, df_kpi, df_data):
    st.subheader("📊 تقارير الموظفين الفردية")
    
    if df_emp is None or df_emp.empty:
        st.warning("⚠️ لا توجد بيانات موظفين.")
        return
    
    # فلترة حسب المقيم (للأدمن العادي)
    role = st.session_state.get("role", "user")
    reviewer_col = df_emp.columns[3] if len(df_emp.columns) > 3 else df_emp.columns[-1]
    
    if role != "super_admin":
        current_reviewer = st.session_state.get("reviewer", "")
        if current_reviewer:
            df_emp = df_emp[df_emp[reviewer_col].astype(str).str.strip() == current_reviewer]
    
    sel_emp = st.selectbox("👤 اختر الموظف", 
                           sorted([str(e).strip() for e in df_emp["EmployeeName"].dropna() 
                                   if str(e).strip() not in ("","nan")]))
    
    if not sel_emp:
        return
    
    sel_year = st.selectbox("📅 السنة", [2025, 2026, 2027], index=0)
    
    if st.button("📊 عرض التقرير", type="primary", use_container_width=True):
        show_employee_report_detail(sel_emp, sel_year, df_emp, df_kpi, df_data)


def show_employee_report_detail(emp_name, year, df_emp, df_kpi, df_data):
    # معلومات الموظف
    emp_row = df_emp[df_emp["EmployeeName"] == emp_name]
    if emp_row.empty:
        st.error("لم يتم العثور على بيانات الموظف.")
        return
    
    emp_row = emp_row.iloc[0]
    job_title = str(emp_row.iloc[1]).strip() if len(emp_row) > 1 else "---"
    dept = str(emp_row.iloc[2]).strip() if len(emp_row) > 2 else "---"
    mgr = str(emp_row.iloc[3]).strip() if len(emp_row) > 3 else "---"
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("👤 الموظف", emp_name)
    col2.metric("💼 الوظيفة", job_title)
    col3.metric("🏭 القسم", dept)
    col4.metric("👨‍💼 المدير", mgr)
    
    st.markdown("---")
    
    # جدول التقييمات الشهرية
    st.markdown("### 📈 ملخص التقييمات الشهرية")
    
    MONTHS_AR = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
                 "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
    MONTHS_EN = ["January", "February", "March", "April", "May", "June",
                 "July", "August", "September", "October", "November", "December"]
    
    def verbal_grade(score):
        if score >= 90: return "ممتاز"
        if score >= 80: return "جيد جداً"
        if score >= 70: return "جيد"
        if score >= 60: return "مقبول"
        return "ضعيف"
    
    report_data = []
    scores = []
    
    for idx, month_en in enumerate(MONTHS_EN):
        month_ar = MONTHS_AR[idx]
        
        m_rows = pd.DataFrame()
        if df_data is not None and not df_data.empty:
            m_rows = df_data[
                (df_data["EmployeeName"] == emp_name) &
                (df_data["Year"] == int(year)) &
                (df_data["Month"] == month_en)
            ]
        
        if not m_rows.empty:
            score = round(float(m_rows["KPI_%"].sum()), 1)
            scores.append(score)
            verbal = verbal_grade(score)
        else:
            score = 0.0
            verbal = "---"
        
        report_data.append({
            "الشهر": month_ar,
            "الدرجة": score,
            "التقييم": verbal,
            "عدد المؤشرات": len(m_rows) if not m_rows.empty else 0,
            "الإجراءات التأديبية": "-"
        })
    
    df_report = pd.DataFrame(report_data)
    
    # عرض الجدول بدون background_gradient لتجنب الخطأ
    st.dataframe(
        df_report.style.format({
            "الدرجة": "{:.1f}",
            "عدد المؤشرات": "{:.0f}"
        }),
        use_container_width=True,
        hide_index=True
    )
    
    avg_score = round(sum(scores) / max(len(scores), 1), 1)
    avg_verbal = verbal_grade(avg_score)
    
    st.markdown("#### 📊 المتوسط السنوي")
    col1, col2, col3 = st.columns(3)
    col1.metric("🎯 متوسط الدرجة", f"{avg_score}%")
    col2.metric("📝 التقييم اللفظي", avg_verbal)
    col3.metric("📅 عدد الأشهر المكتملة", f"{len(scores)}/12")


def download_report_pdf(emp_name, year, df_emp, df_kpi, df_data):
    st.info("📥 ميزة PDF قيد التطوير...")
