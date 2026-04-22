# entry.py
import streamlit as st
import pandas as pd
from constants import MONTHS_AR, MONTHS_EN, MONTH_MAP, PERSONAL_KPIS, PERSONAL_WEIGHT
from data_loader import save_evaluation
from auth import get_current_reviewer, get_current_role
from calculations import calc_kpi_score, rating_label, rating_label_color, verbal_grade, grade_color_hex

# ══════════════════════════════════════════════════════════════════
# استيراد دوال قاعدة البيانات الجديدة
# ══════════════════════════════════════════════════════════════════
try:
    from database_manager import (
        get_previous_evaluation,
        get_employee_disciplinary_actions,
        load_employees_db
    )
    _DB_FUNCS_AVAILABLE = True
except ImportError:
    _DB_FUNCS_AVAILABLE = False
    def get_previous_evaluation(emp_name, year):
        return None
    def get_employee_disciplinary_actions(emp_name, year=None, month=None):
        return []
    def load_employees_db():
        return []


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


def _save_draft(emp, month, year, job_grades, pers_grades, notes, training):
    """حفظ مسودة في session_state."""
    key = _draft_key(emp, month, year)
    st.session_state[key] = {
        "job_grades":   {k: v for k, v in job_grades.items()},
        "pers_grades":  {k: v for k, v in pers_grades.items()},
        "notes":        notes,
        "training":     training,
        "timestamp":    pd.Timestamp.now().strftime("%H:%M"),
    }


def _load_draft(emp, month, year):
    """تحميل المسودة من session_state."""
    return st.session_state.get(_draft_key(emp, month, year), None)


def _clear_draft(emp, month, year):
    key = _draft_key(emp, month, year)
    if key in st.session_state:
        del st.session_state[key]


def _completion_indicator(df_data, emp_list, year):
    """مؤشر اكتمال التقييمات."""
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


# ══════════════════════════════════════════════════════════════════
# دالة عرض نتيجة التقييم السابق
# ══════════════════════════════════════════════════════════════════
def _render_previous_evaluation(emp_name, current_year):
    """عرض نتيجة التقييم السابق للموظف."""
    prev_eval = get_previous_evaluation(emp_name, current_year)
    
    if prev_eval is None:
        # لم يتم تقييمه سابقاً
        st.markdown(f"""
        <div style="background:#FEF3C7;border:2px solid #F59E0B;border-radius:10px;
                    padding:12px 16px;margin:10px 0;">
            <div style="font-size:13px;color:#92400E;font-weight:bold;">
                📋 نتيجة التقييم السابق ({current_year - 1})
            </div>
            <div style="font-size:14px;color:#78350F;margin-top:6px;">
                ⚠️ لم يتم تقييم هذا الموظف سابقاً
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        # يوجد تقييم سابق
        prev_score = prev_eval.get("score", 0)
        prev_verbal = prev_eval.get("verbal", "---")
        prev_year = prev_eval.get("year", current_year - 1)
        
        # تحديد لون النتيجة
        if prev_score >= 90:
            score_color = "#15803d"
            bg_color = "#DCFCE7"
            border_color = "#22C55E"
        elif prev_score >= 80:
            score_color = "#1d4ed8"
            bg_color = "#DBEAFE"
            border_color = "#3B82F6"
        elif prev_score >= 70:
            score_color = "#ca8a04"
            bg_color = "#FEF9C3"
            border_color = "#EAB308"
        elif prev_score >= 60:
            score_color = "#ea580c"
            bg_color = "#FFEDD5"
            border_color = "#F97316"
        else:
            score_color = "#dc2626"
            bg_color = "#FEE2E2"
            border_color = "#EF4444"
        
        st.markdown(f"""
        <div style="background:{bg_color};border:2px solid {border_color};border-radius:10px;
                    padding:12px 16px;margin:10px 0;">
            <div style="font-size:13px;color:#374151;font-weight:bold;">
                📋 نتيجة التقييم السابق ({prev_year})
            </div>
            <div style="display:flex;justify-content:space-around;margin-top:10px;">
                <div style="text-align:center;">
                    <div style="font-size:11px;color:#6B7280;">الدرجة</div>
                    <div style="font-size:1.8rem;font-weight:bold;color:{score_color};">
                        {prev_score:.1f}%
                    </div>
                </div>
                <div style="text-align:center;">
                    <div style="font-size:11px;color:#6B7280;">التقييم اللفظي</div>
                    <div style="font-size:1.2rem;font-weight:bold;color:{score_color};">
                        {prev_verbal}
                    </div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# دالة عرض الإجراءات التأديبية
# ══════════════════════════════════════════════════════════════════
def _render_disciplinary_actions(emp_name, year):
    """عرض الإجراءات التأديبية للموظف."""
    actions = get_employee_disciplinary_actions(emp_name, year)
    
    if not actions:
        st.markdown(f"""
        <div style="background:#DCFCE7;border:2px solid #22C55E;border-radius:10px;
                    padding:10px 14px;margin:10px 0;">
            <div style="font-size:13px;color:#166534;font-weight:bold;">
                ✅ الإجراءات التأديبية: لا يوجد إجراءات تأديبية
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        actions_html = ""
        for action in actions:
            action_date = action.get("action_date", "")
            action_type = action.get("action_type", "")
            description = action.get("description", "")
            actions_html += f"""
            <div style="background:#FEF2F2;padding:8px 12px;border-radius:6px;
                        margin:6px 0;border-right:4px solid #EF4444;">
                <div style="font-size:12px;color:#991B1B;font-weight:bold;">
                    📅 {action_date} — {action_type}
                </div>
                <div style="font-size:11px;color:#7F1D1D;margin-top:4px;">
                    {description}
                </div>
            </div>
            """
        
        st.markdown(f"""
        <div style="background:#FEE2E2;border:2px solid #EF4444;border-radius:10px;
                    padding:12px 16px;margin:10px 0;">
            <div style="font-size:13px;color:#991B1B;font-weight:bold;margin-bottom:8px;">
                ⚠️ الإجراءات التأديبية ({len(actions)} إجراء)
            </div>
            {actions_html}
        </div>
        """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# دالة عرض معلومات الموظف الإضافية
# ══════════════════════════════════════════════════════════════════
def _render_employee_extra_info(emp_name, df_emp):
    """عرض معلومات الموظف الإضافية (تاريخ التعيين، الراتب)."""
    extra_info = {}
    
    # محاولة جلب المعلومات من قاعدة البيانات أولاً
    if _DB_FUNCS_AVAILABLE:
        employees_db = load_employees_db()
        for emp in employees_db:
            if emp.get("EmployeeName") == emp_name:
                extra_info = emp
                break
    
    # إذا لم نجد في DB، نحاول من df_emp
    if not extra_info:
        emp_row = df_emp[df_emp["EmployeeName"] == emp_name]
        if not emp_row.empty:
            row = emp_row.iloc[0]
            extra_info = {
                "hire_date": row.get("hire_date", row.get("تاريخ التعيين", "")),
                "salary": row.get("salary", row.get("الراتب", "")),
            }
    
    hire_date = extra_info.get("hire_date", "")
    salary = extra_info.get("salary", "")
    
    # عرض المعلومات الإضافية فقط إذا كانت موجودة
    if hire_date or salary:
        info_parts = []
        if hire_date and str(hire_date).strip() not in ("", "nan", "None"):
            info_parts.append(f"<b>📅 تاريخ التعيين:</b> {hire_date}")
        if salary and str(salary).strip() not in ("", "nan", "None", "0"):
            info_parts.append(f"<b>💰 الراتب الحالي:</b> {salary}")
        
        if info_parts:
            st.markdown(f"""
            <div style="background:#F0FDF4;padding:10px 14px;border-radius:8px;
                        border-right:4px solid #22C55E;margin:8px 0;font-size:13px;">
                {" &nbsp;|&nbsp; ".join(info_parts)}
            </div>
            """, unsafe_allow_html=True)


def render_entry(df_emp, df_kpi, df_data):
    st.subheader("📋 إدخال التقييم الشهري")
    st.markdown(INPUT_CSS, unsafe_allow_html=True)

    df_data = _safe_df(df_data)

    role             = get_current_role()
    current_reviewer = get_current_reviewer()
    is_super_admin   = (role == "super_admin")
    is_admin         = (role in ("admin", "super_admin"))
    reviewer_col     = df_emp.columns[3] if len(df_emp.columns) > 3 else df_emp.columns[-1]

    r1c1, r1c2, r1c3 = st.columns(3)

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
                st.markdown(f"""<div style="background:#EFF6FF;padding:10px 14px;
                    border-radius:8px;border-right:4px solid #1E3A8A;">
                    <b>👨‍💼 المقيم:</b> {sel_reviewer}</div>""", unsafe_allow_html=True)
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
            st.markdown(f"""<div style="background:#EFF6FF;padding:10px 14px;
                border-radius:8px;border-right:4px solid #1E3A8A;">
                <b>👨‍💼 المقيم:</b> {sel_reviewer}</div>""", unsafe_allow_html=True)

    with r1c2:
        reviewer_chosen = sel_reviewer != "-- اختر المقيم --"
        if not reviewer_chosen:
            emp_list = []
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

    if not is_super_admin and is_admin and not current_reviewer \
            and sel_reviewer == "-- اختر المقيم --":
        st.info("⬆️ اختر المقيم أولاً.")
        return

    # ══════════════════════════════════════════════════════════════
    # مؤشر اكتمال التقييمات
    # ══════════════════════════════════════════════════════════════
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
                st.markdown(f"""
                <div style="background:white;border:1px solid #E2E8F0;border-radius:10px;
                            padding:10px;text-align:center;margin-bottom:8px;">
                    <div style="font-size:11px;color:#64748B;margin-bottom:4px;
                                white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
                        {emp[:20]}
                    </div>
                    <div style="font-size:1.4rem;font-weight:bold;color:{color};">
                        {done_count}/12
                    </div>
                    <div style="background:#F1F5F9;border-radius:4px;height:6px;margin-top:4px;">
                        <div style="background:{color};width:{pct}%;height:6px;border-radius:4px;"></div>
                    </div>
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
    emp_row   = emp_row.iloc[0]
    job_title = str(emp_row.iloc[1]).strip()
    dept_name = str(emp_row.iloc[2]).strip()
    mgr_name  = str(emp_row.iloc[3]).strip()

    if not df_data.empty and "EmployeeName" in df_data.columns:
        dup = df_data[
            (df_data["EmployeeName"] == sel_emp) &
            (df_data["Month"]        == MONTH_MAP.get(sel_month, sel_month)) &
            (df_data["Year"]         == int(sel_year))
        ]
        if not dup.empty:
            st.error(f"⚠️ يوجد تقييم محفوظ لـ ({sel_emp}) في {sel_month} {sel_year}.")
            return

    # ══════════════════════════════════════════════════════════════
    # رأس بيانات الموظف + زر إلغاء
    # ══════════════════════════════════════════════════════════════
    hc1, hc2 = st.columns([5, 1])
    with hc1:
        st.markdown(f"""
        <div style="background:#EFF6FF;padding:12px;border-radius:10px;
            border-right:5px solid #1E3A8A;margin-bottom:14px;">
            <b>👤 الموظف:</b> {sel_emp} &nbsp;|&nbsp;
            <b>💼 الوظيفة:</b> {job_title} &nbsp;|&nbsp;
            <b>🏢 القسم:</b> {dept_name} &nbsp;|&nbsp;
            <b>👨‍💼 المقيم:</b> {mgr_name}
        </div>""", unsafe_allow_html=True)
    with hc2:
        if st.button("❌ إلغاء", use_container_width=True, help="إلغاء والعودة"):
            _clear_draft(sel_emp, sel_month, sel_year)
            st.session_state.pop("sel_emp", None)
            st.rerun()

    # ══════════════════════════════════════════════════════════════
    # 🆕 عرض معلومات الموظف الإضافية
    # ══════════════════════════════════════════════════════════════
    _render_employee_extra_info(sel_emp, df_emp)

    # ══════════════════════════════════════════════════════════════
    # 🆕 عرض نتيجة التقييم السابق
    # ══════════════════════════════════════════════════════════════
    _render_previous_evaluation(sel_emp, sel_year)

    # ══════════════════════════════════════════════════════════════
    # 🆕 عرض الإجراءات التأديبية
    # ══════════════════════════════════════════════════════════════
    _render_disciplinary_actions(sel_emp, sel_year)

    st.markdown("---")

    kpi_rows_raw = df_kpi[df_kpi["JobTitle"].astype(str).str.strip() == job_title]
    if kpi_rows_raw.empty:
        st.warning(f"⚠️ لا توجد مؤشرات KPI لوظيفة '{job_title}'.")
        return

    job_kpis  = kpi_rows_raw[~kpi_rows_raw["KPI_Name"].isin(PERSONAL_KPIS)]
    pers_kpis = kpi_rows_raw[kpi_rows_raw["KPI_Name"].isin(PERSONAL_KPIS)]

    # تحميل المسودة إن وجدت
    draft = _load_draft(sel_emp, sel_month, sel_year)
    if draft:
        st.info(f"📝 يوجد مسودة محفوظة بتاريخ {draft['timestamp']} — يمكنك متابعة الإدخال أو البدء من جديد.")
        if st.button("🗑️ حذف المسودة والبدء من جديد"):
            _clear_draft(sel_emp, sel_month, sel_year)
            st.rerun()

    COLORS = ["#DBEAFE","#E0F2FE","#EDE9FE","#FCE7F3","#D1FAE5",
              "#FEF3C7","#FEE2E2","#F0FDF4","#EFF6FF","#FDF4FF"]

    st.markdown("### 🎯 مؤشرات الأداء الوظيفي")

    job_grades = {}
    for i, (_, row) in enumerate(job_kpis.iterrows()):
        kname  = str(row["KPI_Name"]).strip()
        weight = float(row["Weight"])
        bg     = COLORS[i % len(COLORS)]
        draft_val = draft["job_grades"].get(kname, (weight, 0.0))[1] if draft else 0.0

        col_name, col_inp, col_info = st.columns([4, 1.2, 1.5])
        with col_name:
            st.markdown(f"""
            <div style="background:{bg};padding:10px 14px;border-radius:8px;
                        border-right:4px solid #1E3A8A;height:52px;
                        display:flex;align-items:center;">
                <b style="font-size:13px;color:#1E3A8A;">{kname}</b>
                <span style="margin-right:8px;color:#64748B;font-size:11px;">
                    (الوزن: {weight}%)
                </span>
            </div>""", unsafe_allow_html=True)
        with col_inp:
            val = st.number_input("", min_value=0.0, max_value=float(weight),
                value=float(draft_val), step=0.1, key=f"kpi_{kname}",
                label_visibility="collapsed", format="%.1f")
            job_grades[kname] = (weight, val)
        with col_info:
            score = calc_kpi_score(val, weight)
            lbl   = rating_label(score / weight * 100 if weight else 0)
            clr   = rating_label_color(lbl)
            st.markdown(f"""
            <div style="background:{clr}22;border:1px solid {clr};
                        border-radius:6px;padding:4px 8px;text-align:center;
                        font-size:11px;font-weight:bold;color:{clr};">
                {score:.1f}% — {lbl}
            </div>""", unsafe_allow_html=True)

    job_total = sum(calc_kpi_score(v, w) for w, v in job_grades.values())

    st.markdown("---")
    st.markdown("### 🌟 مؤشرات الصفات الشخصية")

    pers_grades = {}
    pers_source = pers_kpis if not pers_kpis.empty else \
        pd.DataFrame([{"KPI_Name": k, "Weight": PERSONAL_WEIGHT} for k in PERSONAL_KPIS])

    for i, (_, row) in enumerate(pers_source.iterrows()):
        kname  = str(row["KPI_Name"]).strip()
        weight = float(row["Weight"])
        bg     = COLORS[(i+5) % len(COLORS)]
        draft_val2 = draft["pers_grades"].get(kname, (weight, 0.0))[1] if draft else 0.0

        col_name2, col_inp2, col_info2 = st.columns([4, 1.2, 1.5])
        with col_name2:
            st.markdown(f"""
            <div style="background:{bg};padding:10px 14px;border-radius:8px;
                        border-right:4px solid #ED7D31;height:52px;
                        display:flex;align-items:center;">
                <b style="font-size:13px;color:#92400E;">{kname}</b>
                <span style="margin-right:8px;color:#64748B;font-size:11px;">
                    (الوزن: {weight}%)
                </span>
            </div>""", unsafe_allow_html=True)
        with col_inp2:
            val2 = st.number_input("", min_value=0.0, max_value=float(weight),
                value=float(draft_val2), step=0.1, key=f"pers_{kname}",
                label_visibility="collapsed", format="%.1f")
            pers_grades[kname] = (weight, val2)
        with col_info2:
            lbl2 = rating_label(val2 / weight * 100 if weight else 0)
            clr2 = rating_label_color(lbl2)
            st.markdown(f"""
            <div style="background:{clr2}22;border:1px solid {clr2};
                        border-radius:6px;padding:4px 8px;text-align:center;
                        font-size:11px;font-weight:bold;color:{clr2};">
                {val2:.1f}% — {lbl2}
            </div>""", unsafe_allow_html=True)

    pers_total  = sum(v for _, v in pers_grades.values())
    grand_total = job_total + pers_total
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
        notes_val    = draft["notes"] if draft else ""
        notes        = st.text_area("📝 ملاحظات المقيم", value=notes_val, key="notes_inp", height=80)
    with col_t:
        training_val = draft["training"] if draft else ""
        training     = st.text_area("🎓 الاحتياجات التدريبية", value=training_val, key="train_inp", height=80)

    rev_name = sel_reviewer if (sel_reviewer != "-- اختر المقيم --") else mgr_name

    # ══════════════════════════════════════════════════════════════
    # أزرار الحفظ والمسودة والإلغاء
    # ══════════════════════════════════════════════════════════════
    b1, b2, b3 = st.columns(3)

    with b1:
        if st.button("💾 حفظ التقييم", type="primary", use_container_width=True):
            kpi_rows = []
            for kname, (weight, val) in job_grades.items():
                score = calc_kpi_score(val, weight)
                lbl   = rating_label(score / weight * 100 if weight else 0)
                kpi_rows.append((kname, weight, score, lbl))
            for kname, (weight, val) in pers_grades.items():
                lbl = rating_label(val / weight * 100 if weight else 0)
                kpi_rows.append((kname, weight, float(val), lbl))
            ok, err = save_evaluation(
                sel_emp, sel_month, sel_year, rev_name, dept_name,
                kpi_rows, notes, training
            )
            if ok:
                _clear_draft(sel_emp, sel_month, sel_year)
                st.success(f"✅ تم حفظ تقييم {sel_emp} لشهر {sel_month} {sel_year} بنجاح!")
                st.cache_data.clear()
                st.rerun()
            else:
                st.error(f"❌ فشل الحفظ: {err}")

    with b2:
        if st.button("📌 حفظ مسودة", use_container_width=True,
                     help="احفظ التقدم الحالي للعودة إليه لاحقاً"):
            _save_draft(sel_emp, sel_month, sel_year,
                        job_grades, pers_grades, notes, training)
            st.success("✅ تم حفظ المسودة — يمكنك العودة إليها لاحقاً.")

    with b3:
        if st.button("❌ إلغاء", use_container_width=True,
                     help="إلغاء والعودة بدون حفظ"):
            _clear_draft(sel_emp, sel_month, sel_year)
            st.session_state.pop("sel_emp", None)
            st.rerun()
