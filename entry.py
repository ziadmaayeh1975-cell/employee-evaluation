import streamlit as st
import pandas as pd
from constants import MONTHS_AR, MONTHS_EN, MONTH_MAP, PERSONAL_KPIS, PERSONAL_WEIGHT
from data_loader import save_evaluation
from auth import get_current_reviewer, get_current_role
from calculations import rating_label, rating_label_color, verbal_grade, grade_color_hex

# ✅ استيراد نظام الإجراءات التأديبية الجديد (من قاعدة البيانات)
try:
    from disciplinary_manager import (
        load_actions as load_disciplinary_actions,
        get_actions_by_employee as get_employee_disciplinary,
        get_actions_summary as format_disciplinary_text
    )
    DISCIPLINARY_AVAILABLE = True
except ImportError:
    DISCIPLINARY_AVAILABLE = False
    st.warning("⚠️ نظام الإجراءات التأديبية غير متاح")

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

    # مؤشر اكتمال التقييمات
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
    
    # جلب رقم الموظف إذا كان موجوداً في df_emp
    emp_id = ""
    if "رقم الموظف" in df_emp.columns:
        emp_id = str(emp_row.get("رقم الموظف", ""))

    if not df_data.empty and "EmployeeName" in df_data.columns:
        dup = df_data[
            (df_data["EmployeeName"] == sel_emp) &
            (df_data["Month"]        == MONTH_MAP.get(sel_month, sel_month)) &
            (df_data["Year"]         == int(sel_year))
        ]
        if not dup.empty:
            st.error(f"⚠️ يوجد تقييم محفوظ لـ ({sel_emp}) في {sel_month} {sel_year}.")
            return

    # رأس بيانات الموظف + زر إلغاء
    hc1, hc2 = st.columns([5, 1])
    with hc1:
        st.markdown(f"""
        <div style="background:#EFF6FF;padding:12px;border-radius:10px;
            border-right:5px solid #1E3A8A;margin-bottom:14px;">
            <b>👤 الموظف:</b> {sel_emp}<br>
            <b>🆔 رقم الموظف:</b> {emp_id}<br>
            <b>💼 الوظيفة:</b> {job_title}<br>
            <b>🏢 القسم:</b> {dept_name}<br>
            <b>👨‍💼 المقيم:</b> {mgr_name}
        </div>""", unsafe_allow_html=True)
    with hc2:
        if st.button("❌ إلغاء", use_container_width=True, help="إلغاء والعودة"):
            _clear_draft(sel_emp, sel_month, sel_year)
            st.session_state.pop("sel_emp", None)
            st.rerun()

    # ═══════════════════════════════════════════════════════════════════
    # 📋 الإجراءات التأديبية (باستخدام النظام الجديد)
    # ═══════════════════════════════════════════════════════════════════
    if DISCIPLINARY_AVAILABLE:
        try:
            # تحميل جميع الإجراءات من قاعدة البيانات
            all_actions = load_disciplinary_actions()
            if all_actions:
                # تحويل اسم الشهر العربي إلى رقم الشهر (للمقارنة مع قاعدة البيانات)
                month_number = MONTHS_AR.index(sel_month) + 1  # يناير = 1, فبراير = 2, إلخ
                
                # جلب الإجراءات التي تطابق الموظف والسنة والشهر المحدد
                disc_actions = get_employee_disciplinary(sel_emp, sel_year, month_number)
                
                if disc_actions:
                    st.markdown("---")
                    st.markdown("#### ⚠️ الإجراءات التأديبية")
                    st.warning(f"⚠️ يوجد **{len(disc_actions)}** إجراء(ات) تأديبي(ة) مسجلة للموظف في هذا الشهر")
                    
                    # عرض جدول الإجراءات
                    display_data = []
                    for a in disc_actions:
                        display_data.append({
                            "التاريخ": a.get("action_date", ""),
                            "نوع الإنذار": a.get("warning_type", ""),
                            "السبب": a.get("reason", ""),
                            "خصم (أيام)": a.get("deduction_days", 0)
                        })
                    
                    if display_data:
                        st.dataframe(pd.DataFrame(display_data), hide_index=True, use_container_width=True)
                    
                    # تخزين الإجراءات في session_state للتقرير
                    st.session_state[f"disciplinary_{sel_emp}_{sel_month}_{sel_year}"] = disc_actions
                else:
                    st.info("✅ لا توجد إجراءات تأديبية للموظف في هذا الشهر")
        except Exception as e:
            st.error(f"⚠️ خطأ في تحميل الإجراءات التأديبية: {e}")

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

    # ═══════════════════════════════════════════════════════════════════
    # مؤشرات الأداء الوظيفي (المقيم يرى فقط 0-100)
    # ═══════════════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("### 🎯 مؤشرات الأداء الوظيفي")
    st.caption("📌 أدخل النسبة المئوية لكل مؤشر (من 0 إلى 100)")

    job_pct_values = {}
    job_actual_scores = {}
    job_total_weight = 0.0

    for i, (_, row) in enumerate(job_kpis.iterrows()):
        kname  = str(row["KPI_Name"]).strip()
        weight = float(row["Weight"])
        bg     = COLORS[i % len(COLORS)]
        
        draft_val = draft["job_pct_values"].get(kname, 0.0) if draft else 0.0

        col_name, col_inp, col_info = st.columns([4, 1.5, 1.5])
        with col_name:
            st.markdown(f"""
            <div style="background:{bg};padding:10px 14px;border-radius:8px;
                        border-right:4px solid #1E3A8A;height:52px;
                        display:flex;align-items:center;">
                <b style="font-size:13px;color:#1E3A8A;">{kname}</b>
                <span style="margin-right:8px;color:#64748B;font-size:11px;">
                    (الوزن في النظام: {weight}%)
                </span>
            </div>""", unsafe_allow_html=True)
        
        with col_inp:
            user_pct = st.number_input(
                "", 
                min_value=0.0, 
                max_value=100.0, 
                value=float(draft_val), 
                step=1.0, 
                key=f"kpi_{kname}",
                label_visibility="collapsed",
                format="%.0f"
            )
            job_pct_values[kname] = user_pct
            
            actual_score = _calculate_actual_score(user_pct, weight)
            job_actual_scores[kname] = actual_score
            job_total_weight += weight
        
        with col_info:
            if user_pct >= 90:
                lbl, clr = "ممتاز", "#15803d"
            elif user_pct >= 80:
                lbl, clr = "جيد جداً", "#1d4ed8"
            elif user_pct >= 70:
                lbl, clr = "جيد", "#92400e"
            elif user_pct >= 60:
                lbl, clr = "متوسط", "#b45309"
            else:
                lbl, clr = "ضعيف", "#b91c1c"
            
            st.markdown(f"""
            <div style="background:{clr}22;border:1px solid {clr};
                        border-radius:6px;padding:4px 8px;text-align:center;
                        font-size:11px;font-weight:bold;color:{clr};">
                {user_pct:.0f}% — {lbl}
            </div>""", unsafe_allow_html=True)

    job_total = sum(job_actual_scores.values())

    # ═══════════════════════════════════════════════════════════════════
    # مؤشرات الصفات الشخصية (المقيم يرى فقط 0-100)
    # ═══════════════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("### 🌟 مؤشرات الصفات الشخصية")
    st.caption("📌 أدخل النسبة المئوية لكل صفة (من 0 إلى 100)")

    pers_pct_values = {}
    pers_actual_scores = {}
    pers_total_weight = 0.0

    pers_source = pers_kpis if not pers_kpis.empty else \
        pd.DataFrame([{"KPI_Name": k, "Weight": PERSONAL_WEIGHT} for k in PERSONAL_KPIS])

    for i, (_, row) in enumerate(pers_source.iterrows()):
        kname  = str(row["KPI_Name"]).strip()
        weight = float(row["Weight"])
        bg     = COLORS[(i+5) % len(COLORS)]
        
        draft_val2 = draft["pers_pct_values"].get(kname, 0.0) if draft else 0.0

        col_name2, col_inp2, col_info2 = st.columns([4, 1.5, 1.5])
        with col_name2:
            st.markdown(f"""
            <div style="background:{bg};padding:10px 14px;border-radius:8px;
                        border-right:4px solid #ED7D31;height:52px;
                        display:flex;align-items:center;">
                <b style="font-size:13px;color:#92400E;">{kname}</b>
                <span style="margin-right:8px;color:#64748B;font-size:11px;">
                    (الوزن في النظام: {weight}%)
                </span>
            </div>""", unsafe_allow_html=True)
        
        with col_inp2:
            user_pct2 = st.number_input(
                "", 
                min_value=0.0, 
                max_value=100.0, 
                value=float(draft_val2), 
                step=1.0, 
                key=f"pers_{kname}",
                label_visibility="collapsed",
                format="%.0f"
            )
            pers_pct_values[kname] = user_pct2
            
            actual_score2 = _calculate_actual_score(user_pct2, weight)
            pers_actual_scores[kname] = actual_score2
            pers_total_weight += weight
        
        with col_info2:
            if user_pct2 >= 90:
                lbl2, clr2 = "ممتاز", "#15803d"
            elif user_pct2 >= 80:
                lbl2, clr2 = "جيد جداً", "#1d4ed8"
            elif user_pct2 >= 70:
                lbl2, clr2 = "جيد", "#92400e"
            elif user_pct2 >= 60:
                lbl2, clr2 = "متوسط", "#b45309"
            else:
                lbl2, clr2 = "ضعيف", "#b91c1c"
            
            st.markdown(f"""
            <div style="background:{clr2}22;border:1px solid {clr2};
                        border-radius:6px;padding:4px 8px;text-align:center;
                        font-size:11px;font-weight:bold;color:{clr2};">
                {user_pct2:.0f}% — {lbl2}
            </div>""", unsafe_allow_html=True)

    pers_total = sum(pers_actual_scores.values())
    
    grand_total = job_total + pers_total
    verb = verbal_grade(grand_total)
    clr  = grade_color_hex(grand_total)

    st.markdown(f"""
    <div style="background:#F8FAFC;border:1px solid #CBD5E1;border-radius:12px;
                padding:12px;margin:16px 0;direction:rtl;">
        <div style="display:flex;justify-content:space-between;margin-bottom:8px;">
            <span style="color:#475569;">📊 مجموع أوزان المؤشرات الوظيفية:</span>
            <span style="font-weight:bold;">{job_total_weight:.1f} / 80.0</span>
        </div>
        <div style="display:flex;justify-content:space-between;margin-bottom:8px;">
            <span style="color:#475569;">⭐ مجموع أوزان الصفات الشخصية:</span>
            <span style="font-weight:bold;">{pers_total_weight:.1f} / 20.0</span>
        </div>
        <div style="display:flex;justify-content:space-between;">
            <span style="color:#475569;">🎯 الدرجة الفعلية المحققة (وظيفي + شخصي):</span>
            <span style="font-weight:bold;color:{clr};">{grand_total:.1f} / 100.0</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div style="background:white;border:2px solid #1E3A8A;border-radius:12px;
                padding:16px;text-align:center;margin:16px 0;">
        <div style="font-size:12px;color:#64748B;margin-bottom:4px;">
            النتيجة النهائية السنوية المقترحة
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

    b1, b2, b3 = st.columns(3)

    with b1:
        if st.button("💾 حفظ التقييم", type="primary", use_container_width=True):
            kpi_rows = []
            for kname, actual_score in job_actual_scores.items():
                weight = None
                for _, row in job_kpis.iterrows():
                    if str(row["KPI_Name"]).strip() == kname:
                        weight = float(row["Weight"])
                        break
                if weight is not None:
                    pct_for_label = (actual_score / weight) * 100 if weight > 0 else 0
                    lbl = rating_label(pct_for_label)
                else:
                    weight = 0
                    lbl = "جيد"
                kpi_rows.append((kname, weight, actual_score, lbl))
            
            for kname, actual_score in pers_actual_scores.items():
                weight = PERSONAL_WEIGHT
                pct_for_label = (actual_score / weight) * 100 if weight > 0 else 0
                lbl = rating_label(pct_for_label)
                kpi_rows.append((kname, weight, actual_score, lbl))
            
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
                        job_pct_values, pers_pct_values, notes, training)
            st.success("✅ تم حفظ المسودة — يمكنك العودة إليها لاحقاً.")

    with b3:
        if st.button("❌ إلغاء", use_container_width=True,
                     help="إلغاء والعودة بدون حفظ"):
            _clear_draft(sel_emp, sel_month, sel_year)
            st.session_state.pop("sel_emp", None)
            st.rerun()
