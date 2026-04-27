import streamlit as st
import pandas as pd
from constants import MONTHS_AR, MONTHS_EN, MONTH_MAP, PERSONAL_KPIS, PERSONAL_WEIGHT
from data_loader import save_evaluation
from auth import get_current_reviewer, get_current_role
from calculations import rating_label, rating_label_color, verbal_grade, grade_color_hex

# استيراد دوال الإجراءات التأديبية
try:
    from disciplinary_manager import load_actions as load_disciplinary_actions
    from disciplinary_manager import get_actions_by_employee as get_employee_disciplinary
    from disciplinary_manager import get_actions_summary as format_disciplinary_text
    DISCIPLINARY_AVAILABLE = True
except ImportError:
    DISCIPLINARY_AVAILABLE = False

INPUT_CSS = """<style>
div[data-testid="stNumberInput"] input {
    background:#FECACA !important; color:#000 !important;
    font-weight:700 !important; font-size:15px !important;
    text-align:center !important; border:2px solid #EF4444 !important;
    border-radius:8px !important;
}
div[data-testid="stNumberInput"] { justify-content:center !important; }
</style>"""


def _safe_df(df):
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame(columns=["EmployeeName","Month","KPI_Name","Weight",
                                      "KPI_%","Evaluator","Notes","Year","EvalDate","Training"])
    df = df.copy()
    for col in ["EmployeeName","Month","KPI_Name","Weight","KPI_%","Year","EvalDate","Notes","Training"]:
        if col not in df.columns:
            df[col] = pd.Series(dtype="object")
    return df


def _draft_key(emp, month, year):
    return f"draft_{emp}_{month}_{year}"


def _save_draft(emp, month, year, job_pct_values, pers_pct_values, notes, training):
    key = _draft_key(emp, month, year)
    st.session_state[key] = {
        "job_pct_values":   job_pct_values.copy(),
        "pers_pct_values":  pers_pct_values.copy(),
        "notes":            notes,
        "training":         training,
        "timestamp":        pd.Timestamp.now().strftime("%H:%M"),
    }


def _load_draft(emp, month, year):
    return st.session_state.get(_draft_key(emp, month, year), None)


def _clear_draft(emp, month, year):
    key = _draft_key(emp, month, year)
    if key in st.session_state:
        del st.session_state[key]


def _calculate_actual_score(user_pct, weight):
    return (user_pct / 100.0) * weight


def _completion_indicator(df_data, emp_list, year):
    if df_data.empty or "EmployeeName" not in df_data.columns:
        return {}
    result = {}
    for emp in emp_list:
        done = df_data[
            (df_data["EmployeeName"] == emp) &
            (df_data["Year"] == int(year))
        ]["Month"].dropna().unique().tolist()
        result[emp] = done
    return result


def render_entry(df_emp, df_kpi, df_data):
    st.subheader("📋 إدخال التقييم الشهري")
    st.markdown(INPUT_CSS, unsafe_allow_html=True)

    df_data = _safe_df(df_data)

    role             = get_current_role()
    current_reviewer = get_current_reviewer()
    is_super_admin   = (role == "super_admin")
    is_admin         = (role in ("admin", "super_admin"))
    reviewer_col     = "اسم المقيم"  # اسم العمود في df_emp

    r1c1, r1c2, r1c3 = st.columns(3)

    with r1c1:
        if is_super_admin:
            reviewer_list = sorted(df_emp[reviewer_col].dropna().astype(str).str.strip().unique().tolist())
            sel_reviewer = st.selectbox("👨‍💼 اسم المقيم", ["-- اختر المقيم --"] + reviewer_list, key="sel_reviewer")
        elif is_admin:
            if current_reviewer:
                sel_reviewer = current_reviewer
                st.markdown(f"""<div style="background:#EFF6FF;padding:10px 14px;
                    border-radius:8px;border-right:4px solid #1E3A8A;">
                    <b>👨‍💼 المقيم:</b> {sel_reviewer}</div>""", unsafe_allow_html=True)
            else:
                reviewer_list = sorted(df_emp[reviewer_col].dropna().astype(str).str.strip().unique().tolist())
                sel_reviewer = st.selectbox("👨‍💼 اسم المقيم", ["-- اختر المقيم --"] + reviewer_list, key="sel_reviewer")
        else:
            if not current_reviewer:
                st.warning("⚠️ لم يتم ربط حسابك بمقيم. تواصل مع المدير.")
                return
            sel_reviewer = current_reviewer
            st.markdown(f"""<div style="background:#EFF6FF;padding:10px 14px;
                border-radius:8px;border-right:4px solid #1E3A8A;">
                <b>👨‍💼 المقيم:</b> {sel_reviewer}</div>""", unsafe_allow_html=True)

    with r1c2:
        reviewer_chosen = sel_reviewer != "-- اختر المقيم --"
        if not reviewer_chosen:
            emp_list = []
        else:
            emp_list = df_emp[df_emp[reviewer_col].astype(str).str.strip() == sel_reviewer]["EmployeeName"].dropna().astype(str).str.strip().tolist()
            if is_super_admin and not emp_list:
                emp_list = df_emp["EmployeeName"].dropna().astype(str).str.strip().tolist()
        sel_emp = st.selectbox("🎯 اسم الموظف", ["-- اختر --"] + emp_list, key="sel_emp")

    with r1c3:
        sel_year = st.selectbox("🗓️ السنة", [2025, 2026, 2027])

    sel_month = st.selectbox("📅 شهر التقييم", MONTHS_AR)

    if not is_super_admin and is_admin and not current_reviewer and sel_reviewer == "-- اختر المقيم --":
        st.info("⬆️ اختر المقيم أولاً.")
        return

    if emp_list and sel_emp == "-- اختر --":
        st.markdown("---")
        st.markdown("#### 📊 حالة التقييمات لهذا العام")
        completion = _completion_indicator(df_data, emp_list, sel_year)
        total_months = 12
        cols = st.columns(min(4, len(emp_list)))
        for i, emp in enumerate(emp_list[:8]):
            done_count = len(completion.get(emp, []))
            pct = int(done_count / total_months * 100)
            color = "#15803d" if pct == 100 else "#1d4ed8" if pct >= 50 else "#b91c1c"
            with cols[i % len(cols)]:
                st.markdown(f"""<div style="background:white;border:1px solid #E2E8F0;border-radius:10px;padding:10px;text-align:center;margin-bottom:8px;">
                    <div style="font-size:11px;color:#64748B;margin-bottom:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{emp[:20]}</div>
                    <div style="font-size:1.4rem;font-weight:bold;color:{color};">{done_count}/12</div>
                    <div style="background:#F1F5F9;border-radius:4px;height:6px;margin-top:4px;"><div style="background:{color};width:{pct}%;height:6px;border-radius:4px;"></div></div>
                    <div style="font-size:10px;color:{color};margin-top:2px;">{pct}%</div>
                </div>""", unsafe_allow_html=True)
        return

    if sel_emp == "-- اختر --":
        st.info("⬆️ اختر الموظف.")
        return

    emp_row = df_emp[df_emp["EmployeeName"] == sel_emp]
    if emp_row.empty:
        st.warning("⚠️ لم يُعثر على بيانات هذا الموظف.")
        return
    emp_row = emp_row.iloc[0]
    emp_id = str(emp_row.get("رقم الموظف", ""))
    job_title = str(emp_row.get("JobTitle", ""))
    dept_name = str(emp_row.get("القسم", ""))
    mgr_name = str(emp_row.get("اسم المقيم", ""))

    if not df_data.empty and "EmployeeName" in df_data.columns:
        dup = df_data[(df_data["EmployeeName"] == sel_emp) & (df_data["Month"] == MONTH_MAP.get(sel_month, sel_month)) & (df_data["Year"] == int(sel_year))]
        if not dup.empty:
            st.error(f"⚠️ يوجد تقييم محفوظ لـ ({sel_emp}) في {sel_month} {sel_year}.")
            return

    # رأس بيانات الموظف مع رقم الموظف
    hc1, hc2 = st.columns([5, 1])
    with hc1:
        st.markdown(f"""
        <div style="background:#EFF6FF;padding:12px;border-radius:10px;border-right:5px solid #1E3A8A;margin-bottom:14px;">
            <b>👤 الموظف:</b> {sel_emp}<br>
            <b>🆔 رقم الموظف:</b> {emp_id}<br>
            <b>💼 الوظيفة:</b> {job_title}<br>
            <b>🏢 القسم:</b> {dept_name}<br>
            <b>👨‍💼 المقيم:</b> {mgr_name}
        </div>""", unsafe_allow_html=True)
    with hc2:
        if st.button("❌ إلغاء", use_container_width=True):
            _clear_draft(sel_emp, sel_month, sel_year)
            st.session_state.pop("sel_emp", None)
            st.rerun()

    # باقي الكود (مؤشرات الأداء، الإجراءات التأديبية، الحفظ...) كما هو دون تغيير
    # ...
