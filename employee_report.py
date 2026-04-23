import io
from datetime import date
import openpyxl
import pandas as pd
import streamlit as st
import json
import os
try:
    import plotly.graph_objects as go
    PLOTLY_OK = True
except ImportError:
    PLOTLY_OK = False
from constants import MONTHS_AR, MONTHS_EN, MONTHS_SHORT, MONTH_MAP, PERSONAL_KPIS
from calculations import calc_monthly, get_kpi_avgs, verbal_grade, grade_color_hex, kpi_score_to_pct, rating_label
from data_loader import get_emp_notes
from auth import get_current_reviewer, get_current_role
from report_export import build_employee_sheet, print_preview_html

def _reviewer_emp_set(df_emp):
    from auth import get_current_reviewer, get_current_role
    role             = get_current_role()
    current_reviewer = get_current_reviewer()

    if role == "super_admin":
        return None

    if role == "admin" and not current_reviewer:
        return None

    reviewer_col = df_emp.columns[3] if len(df_emp.columns) > 3 else df_emp.columns[-1]
    return set(
        str(e).strip() for e in
        df_emp[df_emp[reviewer_col].astype(str).str.strip() == current_reviewer
               ]["EmployeeName"].dropna().tolist()
        if str(e).strip() not in ("","nan")
    )


def _reviewer_emp_list(df_emp):
    allowed = _reviewer_emp_set(df_emp)
    if allowed is None:
        return df_emp["EmployeeName"].dropna().astype(str).str.strip().tolist()
    return list(allowed)


def _load_evaluations_directly():
    """تحميل التقييمات مباشرة من ملف JSON"""
    file_path = "db/evaluations.json"
    if not os.path.exists(file_path):
        return pd.DataFrame()
    
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            evaluations = data.get("evaluations", [])
        
        if evaluations:
            df = pd.DataFrame(evaluations)
            # توحيد أسماء الأعمدة
            if "KPI_%" in df.columns:
                df["KPI_%"] = pd.to_numeric(df["KPI_%"], errors="coerce").fillna(0)
            if "Weight" in df.columns:
                df["Weight"] = pd.to_numeric(df["Weight"], errors="coerce").fillna(0)
            if "Year" in df.columns:
                df["Year"] = pd.to_numeric(df["Year"], errors="coerce").fillna(2025).astype(int)
            return df
    except:
        pass
    
    return pd.DataFrame()


def render_employee_report(df_emp, df_kpi, df_data):
    st.subheader("📄 نموذج التقييم النهائي للموظف")

    # تحميل التقييمات مباشرة من JSON
    df_data = _load_evaluations_directly()
    
    if df_data.empty:
        st.info("ℹ️ لا توجد تقييمات محفوظة لموظفيك حتى الآن.")
        return

    # فلترة الموظفين
    allowed_emps = set(_reviewer_emp_list(df_emp))

    all_evaluated = sorted([
        str(e) for e in df_data["EmployeeName"].dropna().unique().tolist()
        if str(e).strip() not in ("","nan") and str(e).strip() in allowed_emps
    ])

    if not all_evaluated:
        st.info("ℹ️ لا توجد تقييمات محفوظة لموظفيك حتى الآن.")
        return

    ca, cb, cc = st.columns([2, 2, 1])
    with ca:
        sel2 = st.selectbox("اختر الموظف", all_evaluated, key="rep_emp")
    with cc:
        sel2_year = st.selectbox("السنة", [2025, 2026, 2027], key="rep_year")

    # استخراج الأشهر المتاحة
    emp_data = df_data[
        (df_data["EmployeeName"].astype(str).str.strip() == sel2) &
        (df_data["Year"].astype(int) == int(sel2_year))
    ]
    
    if emp_data.empty:
        st.warning(f"⚠️ لا توجد تقييمات لـ {sel2} في سنة {sel2_year}")
        return

    emp_eval_months_en = emp_data["Month"].dropna().unique().tolist()
    emp_eval_months_ar = [MONTHS_AR[MONTHS_EN.index(m)] for m in emp_eval_months_en if m in MONTHS_EN]

    with cb:
        if emp_eval_months_ar:
            sel2_months = st.multiselect(
                "تصفية بأشهر",
                emp_eval_months_ar, key="rep_months"
            )
        else:
            st.warning("لا توجد تقييمات لهذا الموظف في السنة المختارة.")
            sel2_months = []

    # بيانات الموظف
    ei = df_emp[df_emp["EmployeeName"].astype(str).str.strip() == sel2]
    if ei.empty:
        ei = df_emp.iloc[0]
    else:
        ei = ei.iloc[0]
    
    job2  = str(ei.get("JobTitle", ei.iloc[1] if len(ei) > 1 else "")).strip()
    dept2 = str(ei.get("Department", ei.iloc[2] if len(ei) > 2 else "")).strip()
    mgr2  = str(ei.get("Manager", ei.iloc[3] if len(ei) > 3 else "")).strip()

    # حساب النتائج
    scores = []
    for en in emp_eval_months_en:
        mask = (
            (df_data["EmployeeName"].astype(str).str.strip() == sel2) &
            (df_data["Month"].astype(str).str.strip() == en) &
            (df_data["Year"].astype(int) == int(sel2_year))
        )
        month_data = df_data[mask]
        if not month_data.empty:
            total = month_data["KPI_%"].sum()
            scores.append(total)

    avg_score = round(sum(scores) / len(scores), 1) if scores else 0
    pct2 = avg_score
    verb2 = verbal_grade(pct2)
    clr2 = grade_color_hex(pct2)

    # عرض النتيجة
    st.markdown(f"""
    <div style="background:#F8FAFC;border:1px solid #CBD5E1;border-radius:12px;
                padding:16px;margin-bottom:10px;direction:rtl;">
        <h2 style="margin:0 0 4px;color:#1E3A8A;">{sel2}</h2>
        <p style="margin:3px 0;color:#475569;">💼 {job2} &nbsp;|&nbsp; 🏢 {dept2}
           &nbsp;|&nbsp; 👨‍💼 {mgr2} &nbsp;|&nbsp; 📅 {date.today().strftime('%d/%m/%Y')}</p>
    </div>""", unsafe_allow_html=True)

    st.markdown(f"""
    <div style="background:white;border:2px solid #1E3A8A;border-radius:12px;
                padding:18px;text-align:center;direction:rtl;margin-bottom:12px;">
        <div style="font-size:13px;color:#64748B;font-weight:600;margin-bottom:6px;">
            ✅ النتيجة النهائية
        </div>
        <div style="font-size:3rem;font-weight:bold;color:{clr2};">{pct2}%</div>
        <div style="font-size:1.1rem;color:{clr2};font-weight:600;">{verb2}</div>
    </div>""", unsafe_allow_html=True)

    # جدول التقييمات الشهرية
    st.subheader("📅 التقييم الشهري")
    
    monthly_data = []
    for en in emp_eval_months_en:
        mask = (
            (df_data["EmployeeName"].astype(str).str.strip() == sel2) &
            (df_data["Month"].astype(str).str.strip() == en) &
            (df_data["Year"].astype(int) == int(sel2_year))
        )
        month_total = df_data[mask]["KPI_%"].sum()
        ar_month = MONTHS_AR[MONTHS_EN.index(en)] if en in MONTHS_EN else en
        monthly_data.append({
            "الشهر": ar_month,
            "الدرجة (%)": round(month_total, 1),
            "التقييم اللفظي": verbal_grade(month_total)
        })
    
    if monthly_data:
        st.dataframe(pd.DataFrame(monthly_data), hide_index=True, use_container_width=True)

    # تحميل Excel
    st.subheader("⬇️ تحميل نموذج التقييم")
    st.download_button(
        label="📥 تحميل Excel",
        data=io.BytesIO(),
        file_name=f"تقييم_{sel2}_{date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        disabled=True
    )
    st.caption("خاصية التحميل قيد التطوير")
