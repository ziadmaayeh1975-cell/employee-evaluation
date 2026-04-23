import streamlit as st
from constants import MONTHS_AR, MONTH_MAP, PERSONAL_KPIS, PERSONAL_WEIGHT
from data_loader import save_evaluation
from auth import get_current_reviewer, get_current_role
from calculations import RATING_SCALE, calc_kpi_score, rating_label, rating_label_color

INPUT_CSS = """<style>
div[data-testid="stNumberInput"] input {
    background:#FECACA !important; color:#000 !important;
    font-weight:700 !important; font-size:15px !important;
    text-align:center !important; border:2px solid #EF4444 !important;
    border-radius:8px !important;
}
div[data-testid="stNumberInput"] { justify-content:center !important; }
</style>"""


def render_entry(df_emp, df_kpi, df_data):
    st.subheader("📋 تحديد المقيم والموظف وفترة التقييم")
    st.markdown(INPUT_CSS, unsafe_allow_html=True)

    import pandas as pd
    if df_data is None or not isinstance(df_data, pd.DataFrame):
        df_data = pd.DataFrame(columns=[
            "EmployeeName","Month","KPI_Name","Weight","KPI_%",
            "Evaluator","Notes","Year","EvalDate","Training"
        ])
    else:
        df_data = df_data.copy()
        for col in ["EmployeeName","Month","KPI_Name","Weight","KPI_%","Year","EvalDate","Notes","Training"]:
            if col not in df_data.columns:
                df_data[col] = pd.Series(dtype="object")

    role             = get_current_role()
    current_reviewer = get_current_reviewer()
    is_super_admin   = (role == "super_admin")
    is_admin         = (role in ("admin", "super_admin"))
    reviewer_col     = df_emp.columns[3] if len(df_emp.columns) > 3 else df_emp.columns[-1]

    r1c1, r1c2, r1c3 = st.columns(3)

    # ── اختيار المقيم ───────────────────────────────────────────────
    with r1c1:
        if is_super_admin:
            reviewer_list = sorted([r for r in
                df_emp[reviewer_col].dropna().astype(str).str.strip().unique()
                if r not in ("","nan")])
            sel_reviewer = st.selectbox("👨‍💼 اسم المقيم",
                ["-- اختر المقيم --"] + reviewer_list, key="sel_reviewer")

        elif is_admin:
            if current_reviewer:
                sel_reviewer = current_reviewer
                st.markdown(f"""
                <div style="background:#EFF6FF;padding:10px 14px;
                    border-radius:8px;border-right:4px solid #1E3A8A;">
                    <b>👨‍💼 المقيم:</b> {sel_reviewer}
                </div>""", unsafe_allow_html=True)
            else:
                reviewer_list = sorted([r for r in
                    df_emp[reviewer_col].dropna().astype(str).str.strip().unique()
                    if r not in ("","nan")])
                sel_reviewer = st.selectbox("👨‍💼 اسم المقيم",
                    ["-- اختر المقيم --"] + reviewer_list, key="sel_reviewer")

        else:
            if not current_reviewer:
                st.warning("⚠️ لم يتم ربط حسابك بمقيم. تواصل مع المدير.")
                return
            sel_reviewer = current_reviewer
            st.markdown(f"""
            <div style="background:#EFF6FF;padding:10px 14px;
                border-radius:8px;border-right:4px solid #1E3A8A;">
                <b>👨‍💼 المقيم:</b> {sel_reviewer}
            </div>""", unsafe_allow_html=True)

    # ── اختيار الموظف ───────────────────────────────────────────────
    with r1c2:
        reviewer_chosen = (sel_reviewer != "-- اختر المقيم --") \
            if isinstance(sel_reviewer, str) else True

        if not reviewer_chosen:
            emp_list = []
        elif is_super_admin and sel_reviewer == "-- اختر المقيم --":
            emp_list = sorted([str(e).strip() for e in
                df_emp["EmployeeName"].dropna().tolist()
                if str(e).strip() not in ("","nan")])
        else:
            rev = sel_reviewer
            emp_list = [str(e).strip() for e in
                df_emp[df_emp[reviewer_col].astype(str).str.strip() == rev
                ]["EmployeeName"].dropna().tolist()
                if str(e).strip() not in ("","nan")]

        if is_super_admin and not emp_list:
            emp_list = sorted([str(e).strip() for e in
                df_emp["EmployeeName"].dropna().tolist()
                if str(e).strip() not in ("","nan")])

        sel_emp = st.selectbox("🎯 اسم الموظف",
            ["-- اختر --"] + emp_list, key="sel_emp")

    with r1c3:
        sel_year = st.selectbox("🗓️ السنة", [2025, 2026, 2027])

    sel_month = st.selectbox("📅 شهر التقييم", MONTHS_AR)

    # ── تحقق الاختيار ───────────────────────────────────────────────
    if not is_super_admin and is_admin and not current_reviewer \
            and sel_reviewer == "-- اختر المقيم --":
        st.info("⬆️ اختر المقيم أولاً.")
        return
    if sel_emp == "-- اختر --":
        st.info("⬆️ اختر الموظف.")
        return

    # ── بيانات الموظف ───────────────────────────────────────────────
    emp_row = df_emp[df_emp["EmployeeName"] == sel_emp]
    if emp_row.empty:
        st.warning("⚠️ لم يُعثر على بيانات هذا الموظف.")
        return
    emp_row   = emp_row.iloc[0]
    job_title = str(emp_row.iloc[1]).strip()
    dept_name = str(emp_row.iloc[2]).strip()
    mgr_name  = str(emp_row.iloc[3]).strip()

    # ── تحقق من تكرار التقييم ───────────────────────────────────────
    if not df_data.empty and "EmployeeName" in df_data.columns:
        dup = df_data[
            (df_data["EmployeeName"] == sel_emp) &
            (df_data["Month"]        == MONTH_MAP.get(sel_month, sel_month)) &
            (df_data["Year"]         == int(sel_year))
        ]
        if not dup.empty:
            st.error(f"⚠️ يوجد تقييم محفوظ لـ ({sel_emp}) في {sel_month} {sel_year}.")
            return

    st.markdown(f"""
    <div style="background:#EFF6FF;padding:12px;border-radius:10px;
        border-right:5px solid #1E3A8A;margin-bottom:14px;">
        <b>👤 الموظف:</b> {sel_emp} &nbsp;|&nbsp;
        <b>💼 الوظيفة:</b> {job_title} &nbsp;|&nbsp;
        <b>🏢 القسم:</b> {dept_name} &nbsp;|&nbsp;
        <b>👨‍💼 المقيم:</b> {mgr_name}
    </div>""", unsafe_allow_html=True)

    kpi_rows_raw = df_kpi[df_kpi["JobTitle"].astype(str).str.strip() == job_title]
    if kpi_rows_raw.empty:
        st.warning(f"⚠️ لا توجد مؤشرات KPI لوظيفة '{job_title}'.")
        return

    job_kpis  = kpi_rows_raw[~kpi_rows_raw["KPI_Name"].isin(PERSONAL_KPIS)]
    pers_kpis = kpi_rows_raw[kpi_rows_raw["KPI_Name"].isin(PERSONAL_KPIS)]

    COLORS = ["#DBEAFE","#E0F2FE","#EDE9FE","#FCE7F3","#D1FAE5",
              "#FEF3C7","#FEE2E2","#F0FDF4","#EFF6FF","#FDF4FF"]

    # ── تنبيه طريقة الإدخال ──────────────────────────────────────
    st.info("📌 **طريقة الإدخال:** أدخل درجة من **0 إلى 100** لكل مؤشر. النظام سيحسب النتيجة = (الدرجة × الوزن) ÷ 100")

    # ── مؤشرات الأداء الوظيفي ───────────────────────────────────────
    st.markdown("---")
    st.markdown("### 🎯 مؤشرات الأداء الوظيفي (80%)")

    job_grades = {}
    for i, (_, row) in enumerate(job_kpis.iterrows()):
        kname  = str(row["KPI_Name"]).strip()
        weight = float(row["Weight"])
        bg     = COLORS[i % len(COLORS)]

        col_name, col_inp, col_score, col_lbl = st.columns([3, 1, 1, 1])
        with col_name:
            st.markdown(f"""
            <div style="background:{bg};padding:10px 14px;border-radius:8px;
                        border-right:4px solid #1E3A8A;min-height:52px;
                        display:flex;align-items:center;">
                <b style="font-size:13px;color:#1E3A8A;">{kname}</b>
                <span style="margin-right:8px;color:#64748B;font-size:11px;">
                    (وزن: {weight}%)
                </span>
            </div>""", unsafe_allow_html=True)
        with col_inp:
            val = st.number_input("الدرجة", min_value=0, max_value=100,
                value=0, step=1, key=f"kpi_{kname}", label_visibility="visible")
            job_grades[kname] = (weight, val)
        with col_score:
            score_val = (val * weight) / 100
            st.markdown(f"""
            <div style="background:#F8FAFC;border:1px solid #CBD5E1;
                        border-radius:6px;padding:8px 6px;text-align:center;
                        font-size:13px;font-weight:bold;color:#1E3A8A;margin-top:22px;">
                {score_val:.1f}/{weight}
            </div>""", unsafe_allow_html=True)
        with col_lbl:
            lbl = rating_label(val)
            clr = rating_label_color(lbl)
            st.markdown(f"""
            <div style="background:{clr}22;border:1px solid {clr};
                        border-radius:6px;padding:8px 6px;text-align:center;
                        font-size:12px;font-weight:bold;color:{clr};margin-top:22px;">
                {lbl}
            </div>""", unsafe_allow_html=True)

    # ── مؤشرات الصفات الشخصية ───────────────────────────────────────
    st.markdown("---")
    st.markdown("### 🌟 مؤشرات الصفات الشخصية (20%)")

    pers_grades = {}
    pers_source = pers_kpis if not pers_kpis.empty else \
        pd.DataFrame([{"KPI_Name": k, "Weight": PERSONAL_WEIGHT} for k in PERSONAL_KPIS])

    for i, (_, row) in enumerate(pers_source.iterrows()):
        kname  = str(row["KPI_Name"]).strip()
        weight = float(row["Weight"])
        bg     = COLORS[(i+5) % len(COLORS)]

        col_name2, col_inp2, col_score2, col_lbl2 = st.columns([3, 1, 1, 1])
        with col_name2:
            st.markdown(f"""
            <div style="background:{bg};padding:10px 14px;border-radius:8px;
                        border-right:4px solid #ED7D31;min-height:52px;
                        display:flex;align-items:center;">
                <b style="font-size:13px;color:#92400E;">{kname}</b>
                <span style="margin-right:8px;color:#64748B;font-size:11px;">
                    (وزن: {weight}%)
                </span>
            </div>""", unsafe_allow_html=True)
        with col_inp2:
            val2 = st.number_input("الدرجة", min_value=0, max_value=100,
                value=0, step=1, key=f"pers_{kname}", label_visibility="visible")
            pers_grades[kname] = (weight, val2)
        with col_score2:
            score_val2 = (val2 * weight) / 100
            st.markdown(f"""
            <div style="background:#F8FAFC;border:1px solid #CBD5E1;
                        border-radius:6px;padding:8px 6px;text-align:center;
                        font-size:13px;font-weight:bold;color:#92400E;margin-top:22px;">
                {score_val2:.1f}/{weight}
            </div>""", unsafe_allow_html=True)
        with col_lbl2:
            lbl2 = rating_label(val2)
            clr2 = rating_label_color(lbl2)
            st.markdown(f"""
            <div style="background:{clr2}22;border:1px solid {clr2};
                        border-radius:6px;padding:8px 6px;text-align:center;
                        font-size:12px;font-weight:bold;color:{clr2};margin-top:22px;">
                {lbl2}
            </div>""", unsafe_allow_html=True)

    job_total  = sum((v * w) / 100 for w, v in job_grades.values())
    pers_total = sum((v * w) / 100 for w, v in pers_grades.values())
    grand_total = job_total + pers_total

    from calculations import verbal_grade, grade_color_hex
    verb = verbal_grade(grand_total)
    clr  = grade_color_hex(grand_total)

    st.markdown(f"""
    <div style="background:white;border:2px solid #1E3A8A;border-radius:12px;
                padding:16px;text-align:center;margin:16px 0;">
        <div style="font-size:12px;color:#64748B;margin-bottom:4px;">
            إجمالي النتيجة (وظيفي {job_total:.1f}% + شخصية {pers_total:.1f}%)
        </div>
        <div style="font-size:2.5rem;font-weight:bold;color:{clr};">{grand_total:.1f}%</div>
        <div style="font-size:1rem;color:{clr};font-weight:600;">{verb}</div>
    </div>""", unsafe_allow_html=True)

    st.markdown("---")
    col_n, col_t = st.columns(2)
    with col_n:
        notes    = st.text_area("📝 ملاحظات المقيم", key="notes_inp", height=80)
    with col_t:
        training = st.text_area("🎓 الاحتياجات التدريبية", key="train_inp", height=80)

    rev_name = sel_reviewer if (sel_reviewer != "-- اختر المقيم --") else mgr_name

    if st.button("💾 حفظ التقييم", type="primary", use_container_width=True):
        kpi_rows = []
        for kname, (weight, val) in job_grades.items():
            score = (val * weight) / 100
            lbl   = rating_label(val)
            kpi_rows.append((kname, weight, round(score, 2), lbl))
        for kname, (weight, val) in pers_grades.items():
            score = (val * weight) / 100
            lbl   = rating_label(val)
            kpi_rows.append((kname, weight, round(score, 2), lbl))

        ok, err = save_evaluation(
            sel_emp, sel_month, sel_year, rev_name, dept_name,
            kpi_rows, notes, training
        )
        if ok:
            st.success(f"✅ تم حفظ تقييم {sel_emp} لشهر {sel_month} {sel_year} بنجاح!")
            st.cache_data.clear()
            st.rerun()
        else:
            st.error(f"❌ فشل الحفظ: {err}")

import pandas as pd
