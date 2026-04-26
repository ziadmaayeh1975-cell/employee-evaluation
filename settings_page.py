import os
from datetime import datetime, date
import streamlit as st
from constants import LOGO_PATH
from auth import (hash_pw, load_users, save_users, add_user, update_user, delete_user,
                  load_trial_users, save_trial_users,
                  load_app_settings, save_app_settings)
try:
    from employees_kpis_panel import render_employees_panel, render_kpis_panel
    _EMP_KPI_OK = True
except ImportError:
    _EMP_KPI_OK = False

try:
    from db_settings_panel import render_db_panel
    _DB_PANEL_OK = True
except ImportError:
    _DB_PANEL_OK = False

try:
    from employees_module import render_employee_management, render_cv_reports, load_profiles
except ImportError:
    def render_employee_management(*args, **kwargs):
        st.warning("ملف employees_module.py غير موجود.")
    def render_cv_reports(*args, **kwargs):
        st.warning("ملف employees_module.py غير موجود.")
    def load_profiles():
        return []


# ── دوال مساعدة للأدوار ──────────────────────────────────────────────
def _is_super_admin():
    """الأدمن الرئيسي: role == 'super_admin'"""
    return st.session_state.get("role") == "super_admin"

def _is_admin():
    """أدمن عادي أو رئيسي"""
    return st.session_state.get("role") in ("admin", "super_admin")

def _role_label(role):
    return {"super_admin": "🔴 أدمن رئيسي",
            "admin":       "🟠 أدمن",
            "user":        "🟢 مستخدم"}.get(role, "🟢 مستخدم")

def _role_color(role):
    return {"super_admin": "#991B1B",
            "admin":       "#92400E",
            "user":        "#166534"}.get(role, "#166534")


def render_settings(df_emp, df_kpi, df_data):
    # ── تحقق الصلاحية: admin عادي أو رئيسي ─────────────────────────
    if not _is_admin():
        st.warning("🔒 هذه الصفحة متاحة للمدير فقط.")
        return

    USERS      = load_users()
    TRIAL_DATA = load_trial_users()
    APP_CFG    = load_app_settings()

    _tabs = ["👥 إدارة المستخدمين", "⏳ المستخدمون التجريبيون",
             "🏢 إعدادات الشركة", "👨‍💼 إدارة الموظفين",
             "📊 قائمة مؤشرات الأداء", "📋 ملفات الموظفين (CV)"]
    if _DB_PANEL_OK:
        _tabs.append("🗄️ قاعدة البيانات")
    _tab_objs   = st.tabs(_tabs)
    set_tab1    = _tab_objs[0]
    set_tab2    = _tab_objs[1]
    set_tab3    = _tab_objs[2]
    set_tab4    = _tab_objs[3]
    set_tab_kpi = _tab_objs[4]
    set_tab5    = _tab_objs[5]
    set_tab6    = _tab_objs[6] if _DB_PANEL_OK else None

    # ══════════════════════════════════════════════════════════════════
    # تاب 1 — إدارة المستخدمين
    # ══════════════════════════════════════════════════════════════════
    with set_tab1:
        st.markdown("### 👥 المستخدمون الفعّالون")

        # ── إشعار للأدمن العادي ─────────────────────────────────────
        if not _is_super_admin():
            st.info("👁️ أنت تعرض قائمة المستخدمين فقط — صلاحية التعديل والحذف للأدمن الرئيسي.")

        if "settings_action" not in st.session_state:
            st.session_state.settings_action = None
        if "settings_target" not in st.session_state:
            st.session_state.settings_target = None

        # ── رأس الجدول ───────────────────────────────────────────────
        h1, h2, h3, h4, h5, _ = st.columns([3, 3, 2, 1.2, 1.2, 0.1])
        for col, label in zip([h1, h2, h3, h4, h5],
                               ["الاسم الظاهر", "اسم الدخول", "الدور", "تعديل", "حذف"]):
            col.markdown(
                f"<div style='background:#1E3A8A;color:white;padding:7px 10px;"
                f"border-radius:6px;font-weight:bold;font-size:13px;"
                f"text-align:center;'>{label}</div>", unsafe_allow_html=True)

        st.markdown("<div style='margin:4px 0;'></div>", unsafe_allow_html=True)

        for idx, (uname, udata) in enumerate(USERS.items()):
            role     = udata.get("role", "user")
            role_lbl = _role_label(role)
            role_clr = _role_color(role)
            rbg      = "#F8FAFF" if idx % 2 == 0 else "white"
            c1, c2, c3, c4, c5, c6 = st.columns([3, 3, 2, 1.2, 1.2, 0.1])

            c1.markdown(
                f"<div style='background:{rbg};padding:8px 10px;border-radius:6px;"
                f"font-size:14px;border:1px solid #E2E8F0;'>"
                f"👤 {udata.get('display', uname)}</div>", unsafe_allow_html=True)
            c2.markdown(
                f"<div style='background:{rbg};padding:8px 10px;border-radius:6px;"
                f"font-size:13px;font-family:monospace;color:#1E3A8A;font-weight:600;"
                f"border:1px solid #E2E8F0;'>🔐 {uname}</div>", unsafe_allow_html=True)
            c3.markdown(
                f"<div style='background:{rbg};padding:8px 10px;border-radius:6px;"
                f"font-size:13px;text-align:center;color:{role_clr};font-weight:bold;"
                f"border:1px solid #E2E8F0;'>{role_lbl}</div>", unsafe_allow_html=True)

            # ── أزرار التعديل والحذف — للأدمن الرئيسي فقط ──────────
            with c4:
                if _is_super_admin():
                    if st.button("✏️", key=f"edit_btn_{uname}",
                                 use_container_width=True, help="تعديل المستخدم"):
                        st.session_state.settings_action = "edit"
                        st.session_state.settings_target = uname
                        st.rerun()
                else:
                    st.markdown("<div style='padding:8px;text-align:center;"
                                "color:#CBD5E1;font-size:12px;'>—</div>",
                                unsafe_allow_html=True)

            with c5:
                # لا يمكن حذف المستخدم admin الأساسي، وفقط super_admin يحذف
                if _is_super_admin() and uname != "admin":
                    if st.button("🗑️", key=f"del_btn_{uname}",
                                 use_container_width=True, help="حذف المستخدم"):
                        st.session_state.settings_action = "delete"
                        st.session_state.settings_target = uname
                        st.rerun()
                else:
                    st.markdown("<div style='padding:8px;text-align:center;"
                                "color:#CBD5E1;font-size:12px;'>—</div>",
                                unsafe_allow_html=True)

        st.markdown("---")
        action = st.session_state.settings_action
        target = st.session_state.settings_target

        # ── تأكيد الحذف ─────────────────────────────────────────────
        if action == "delete" and target and _is_super_admin():
            st.error(f"⚠️ هل تريد حذف المستخدم **{target}** "
                     f"({USERS[target].get('display', '')})?")
            cd1, cd2 = st.columns(2)
            with cd1:
                if st.button("✅ تأكيد الحذف", use_container_width=True, type="primary"):
                    del USERS[target]
                    save_users(USERS)
                    st.session_state.settings_action = None
                    st.session_state.settings_target = None
                    st.success("✅ تم الحذف.")
                    st.rerun()
            with cd2:
                if st.button("❌ إلغاء", use_container_width=True):
                    st.session_state.settings_action = None
                    st.session_state.settings_target = None
                    st.rerun()

        # ── تعديل المستخدم ───────────────────────────────────────────
        elif action == "edit" and target and _is_super_admin():
            st.markdown(f"#### ✏️ تعديل: **{USERS[target].get('display', target)}**")
            with st.form("edit_user_form"):
                e1, e2 = st.columns(2)
                with e1:
                    new_disp_e = st.text_input("الاسم الظاهر",
                                                value=USERS[target].get("display", ""))
                    new_uname_e = st.text_input("اسم الدخول (username)",
                                                 value=target,
                                                 help="يمكن تغيير اسم الدخول — سيتم إنشاء مستخدم جديد بالاسم الجديد")
                    new_role_e = st.selectbox(
                        "الدور",
                        ["user", "admin", "super_admin"],
                        index=["user","admin","super_admin"].index(
                            USERS[target].get("role","user"))
                        if USERS[target].get("role","user") in ["user","admin","super_admin"] else 0,
                        format_func=lambda r: {"user":"🟢 مستخدم",
                                               "admin":"🟠 أدمن",
                                               "super_admin":"🔴 أدمن رئيسي"}[r]
                    )
                    new_reviewer_e = st.text_input(
                        "اسم المقيم (كما في EMPLOYEES)",
                        value=USERS[target].get("reviewer", ""),
                        help="يُربط هذا الحساب بالموظفين الذين يقيمهم هذا المقيم"
                    )
                with e2:
                    new_pw_e = st.text_input("كلمة مرور جديدة (اتركها فارغة للإبقاء)",
                                              type="password")

                c1e, c2e = st.columns(2)
                with c1e:
                    save_edit = st.form_submit_button("💾 حفظ التعديلات",
                                                       use_container_width=True)
                with c2e:
                    cancel_edit = st.form_submit_button("❌ إلغاء",
                                                         use_container_width=True)

                if save_edit:
                    new_uname_final = new_uname_e.strip() or target
                    updated_data = {
                        "display":  new_disp_e or target,
                        "role":     new_role_e,
                        "reviewer": new_reviewer_e.strip(),
                        "password": USERS[target]["password"],
                    }
                    if new_pw_e.strip():
                        updated_data["password"] = hash_pw(new_pw_e)

                    # إذا تغير اسم الدخول: احذف القديم وأضف الجديد
                    if new_uname_final != target:
                        del USERS[target]
                    USERS[new_uname_final] = updated_data
                    save_users(USERS)
                    st.session_state.settings_action = None
                    st.session_state.settings_target = None
                    st.success("✅ تم حفظ التعديلات.")
                    st.rerun()

                if cancel_edit:
                    st.session_state.settings_action = None
                    st.session_state.settings_target = None
                    st.rerun()

        # ── إضافة مستخدم جديد — للأدمن الرئيسي فقط ─────────────────
        elif _is_super_admin():
            st.markdown("#### ➕ إضافة مستخدم جديد")
            with st.form("add_user_form"):
                a1, a2 = st.columns(2)
                with a1:
                    new_uname    = st.text_input("اسم الدخول (بالإنجليزي)")
                    new_disp     = st.text_input("الاسم الظاهر (عربي)")
                    new_reviewer = st.text_input(
                        "اسم المقيم (كما في EMPLOYEES)",
                        help="اتركه فارغاً إذا كان admin كامل الصلاحيات"
                    )
                with a2:
                    new_pw   = st.text_input("كلمة المرور", type="password")
                    new_role = st.selectbox(
                        "الدور",
                        ["user", "admin", "super_admin"],
                        format_func=lambda r: {"user":"🟢 مستخدم",
                                               "admin":"🟠 أدمن",
                                               "super_admin":"🔴 أدمن رئيسي"}[r]
                    )

                if st.form_submit_button("➕ إضافة مستخدم", use_container_width=True):
                    if new_uname.strip() and new_pw.strip():
                        if new_uname.strip() in USERS:
                            st.error("⚠️ اسم المستخدم موجود مسبقاً.")
                        else:
                            USERS[new_uname.strip()] = {
                                "password": hash_pw(new_pw),
                                "role":     new_role,
                                "display":  new_disp or new_uname,
                                "reviewer": new_reviewer.strip(),
                            }
                            save_users(USERS)
                            st.success(f"✅ تم إضافة: {new_uname}")
                            st.rerun()
                    else:
                        st.error("⚠️ أدخل اسم الدخول وكلمة المرور.")

    # ══════════════════════════════════════════════════════════════════
    # تاب 2 — المستخدمون التجريبيون (super_admin فقط يعدل/يحذف)
    # ══════════════════════════════════════════════════════════════════
    with set_tab2:
        st.markdown("#### ⏳ المستخدمون التجريبيون")
        if not _is_super_admin():
            st.info("👁️ عرض فقط — صلاحية التعديل والحذف للأدمن الرئيسي.")

        TRIAL_DATA = load_trial_users()

        if TRIAL_DATA:
            now = datetime.now()
            th1,th2,th3,th4,th5,th6 = st.columns([2,2,2,2,1,1])
            for col, lbl in zip([th1,th2,th3,th4,th5,th6],
                                  ["اسم المستخدم","تاريخ البداية","تاريخ الانتهاء",
                                   "الحالة","تعديل","حذف"]):
                col.markdown(
                    f"<div style='background:#7C3AED;color:white;padding:7px 10px;"
                    f"border-radius:6px;font-weight:bold;font-size:13px;"
                    f"text-align:center;'>{lbl}</div>", unsafe_allow_html=True)
            st.markdown("<div style='margin:4px 0'></div>", unsafe_allow_html=True)

            for idx, (tuname, tdata) in enumerate(TRIAL_DATA.items()):
                rbg = "#FAF5FF" if idx%2==0 else "white"
                try:
                    end_dt = datetime.strptime(tdata["end"], "%Y-%m-%d %H:%M")
                    status = "🟢 نشط" if now <= end_dt else "🔴 منتهي"
                    sta_clr = "#15803d" if now <= end_dt else "#b91c1c"
                except:
                    status, sta_clr = "❓", "#64748B"

                tc1,tc2,tc3,tc4,tc5,tc6 = st.columns([2,2,2,2,1,1])
                for col, val in zip([tc1,tc2,tc3],
                                     [tuname, tdata.get("start",""), tdata.get("end","")]):
                    col.markdown(
                        f"<div style='background:{rbg};padding:7px 10px;"
                        f"border-radius:6px;font-size:13px;font-family:monospace;"
                        f"border:1px solid #E2E8F0;'>{val}</div>", unsafe_allow_html=True)
                tc4.markdown(
                    f"<div style='background:{rbg};padding:7px 10px;"
                    f"border-radius:6px;font-size:13px;text-align:center;"
                    f"color:{sta_clr};font-weight:bold;"
                    f"border:1px solid #E2E8F0;'>{status}</div>", unsafe_allow_html=True)

                with tc5:
                    if _is_super_admin():
                        if st.button("✏️", key=f"trial_edit_{tuname}", use_container_width=True):
                            st.session_state.settings_action = "trial_edit"
                            st.session_state.settings_target = tuname
                            st.rerun()
                    else:
                        st.markdown("<div style='padding:8px;text-align:center;color:#CBD5E1;'>—</div>",
                                    unsafe_allow_html=True)
                with tc6:
                    if _is_super_admin():
                        if st.button("🗑️", key=f"trial_del_{tuname}", use_container_width=True):
                            st.session_state.settings_action = "trial_delete"
                            st.session_state.settings_target = tuname
                            st.rerun()
                    else:
                        st.markdown("<div style='padding:8px;text-align:center;color:#CBD5E1;'>—</div>",
                                    unsafe_allow_html=True)
        else:
            st.info("لا يوجد مستخدمون تجريبيون حتى الآن.")

        st.markdown("---")
        action = st.session_state.get("settings_action")
        target = st.session_state.get("settings_target")

        if action == "trial_delete" and target and _is_super_admin():
            st.error(f"حذف المستخدم التجريبي: **{target}**؟")
            cd1, cd2 = st.columns(2)
            with cd1:
                if st.button("✅ تأكيد الحذف", key="conf_trial_del",
                             use_container_width=True, type="primary"):
                    TRIAL_DATA2 = load_trial_users()
                    if target in TRIAL_DATA2:
                        del TRIAL_DATA2[target]; save_trial_users(TRIAL_DATA2)
                    USERS2 = load_users()
                    if target in USERS2:
                        del USERS2[target]; save_users(USERS2)
                    st.session_state.settings_action = None
                    st.session_state.settings_target = None
                    st.success("✅ تم الحذف."); st.rerun()
            with cd2:
                if st.button("❌ إلغاء", key="cancel_trial_del", use_container_width=True):
                    st.session_state.settings_action = None
                    st.session_state.settings_target = None
                    st.rerun()

        elif _is_super_admin() and action in ("trial_edit", None):
            is_edit = (action == "trial_edit" and target)
            tdata   = TRIAL_DATA.get(target, {}) if is_edit else {}
            lbl     = f"✏️ تعديل: {target}" if is_edit else "➕ إضافة مستخدم تجريبي"
            st.markdown(f"#### {lbl}")

            with st.form("trial_form"):
                f1, f2 = st.columns(2)
                with f1:
                    t_uname = st.text_input("اسم المستخدم", value=target or "", disabled=is_edit)
                    t_disp  = st.text_input("الاسم الظاهر", value=tdata.get("display", target or ""))
                    t_pw    = st.text_input("كلمة المرور" + (" (فارغ=لا تغيير)" if is_edit else ""),
                                             type="password")
                with f2:
                    t_start_d = st.date_input("تاريخ البداية",
                        value=datetime.strptime(tdata["start"][:10],"%Y-%m-%d").date()
                              if tdata.get("start") else date.today())
                    t_start_h = st.time_input("ساعة البداية",
                        value=datetime.strptime(tdata["start"],"%Y-%m-%d %H:%M").time()
                              if tdata.get("start")
                              else datetime.now().replace(second=0,microsecond=0).time())
                    t_end_d = st.date_input("تاريخ الانتهاء",
                        value=datetime.strptime(tdata["end"][:10],"%Y-%m-%d").date()
                              if tdata.get("end") else date.today())
                    t_end_h = st.time_input("ساعة الانتهاء",
                        value=datetime.strptime(tdata["end"],"%Y-%m-%d %H:%M").time()
                              if tdata.get("end")
                              else datetime.strptime("23:59","%H:%M").time())

                cf1, cf2 = st.columns(2)
                with cf1:
                    submitted = st.form_submit_button(
                        "💾 حفظ التعديلات" if is_edit else "➕ إنشاء المستخدم",
                        use_container_width=True)
                with cf2:
                    cancelled = st.form_submit_button("❌ إلغاء", use_container_width=True)

                if submitted:
                    uname_final = target if is_edit else t_uname.strip()
                    if not uname_final:
                        st.error("أدخل اسم المستخدم.")
                    else:
                        start_str = f"{t_start_d} {t_start_h.strftime('%H:%M')}"
                        end_str   = f"{t_end_d} {t_end_h.strftime('%H:%M')}"
                        TRIAL_DATA2 = load_trial_users()
                        TRIAL_DATA2[uname_final] = {
                            "display": t_disp or uname_final,
                            "start":   start_str, "end": end_str,
                        }
                        save_trial_users(TRIAL_DATA2)
                        USERS3 = load_users()
                        if uname_final not in USERS3 or not is_edit:
                            USERS3[uname_final] = {
                                "password": hash_pw(t_pw) if t_pw.strip() else hash_pw("Trial123"),
                                "role":     "user",
                                "display":  t_disp or uname_final,
                                "reviewer": "",
                            }
                        elif t_pw.strip():
                            USERS3[uname_final]["password"] = hash_pw(t_pw)
                            USERS3[uname_final]["display"]  = t_disp or uname_final
                        save_users(USERS3)
                        st.session_state.settings_action = None
                        st.session_state.settings_target = None
                        st.success(f"✅ تم {'التعديل' if is_edit else 'الإنشاء'}: {uname_final}")
                        st.rerun()
                if cancelled:
                    st.session_state.settings_action = None
                    st.session_state.settings_target = None
                    st.rerun()

    # ══════════════════════════════════════════════════════════════════
    # تاب 3 — إعدادات الشركة
    # ══════════════════════════════════════════════════════════════════
    with set_tab3:
        st.markdown("#### 🏢 إعدادات الشركة والتقارير")
        APP_CFG = load_app_settings()

        with st.form("company_settings_form"):
            cs1, cs2 = st.columns(2)
            with cs1:
                s_company = st.text_input("🏷️ اسم الشركة / المجموعة",
                    value=APP_CFG.get("company_name",""), placeholder="مثال: مجموعة شركات فنون")
                s_branch  = st.text_input("🏬 اسم الفرع (اختياري)",
                    value=APP_CFG.get("branch_name",""), placeholder="مثال: فرع الرياض")
                s_emp_cnt = st.text_input("👥 عدد الموظفين", value=APP_CFG.get("employee_count",""))
                s_phone   = st.text_input("📞 رقم التواصل", value=APP_CFG.get("contact_phone",""))
            with cs2:
                s_email     = st.text_input("📧 البريد الإلكتروني", value=APP_CFG.get("contact_email",""))
                s_show_logo = st.checkbox("إظهار شعار الشركة في التقارير",
                                          value=APP_CFG.get("show_logo", True))
                s_logo_file = st.file_uploader("🖼️ رفع شعار الشركة (PNG/JPG)",
                                                type=["png","jpg","jpeg"])
                _current_logo = APP_CFG.get("logo_path", LOGO_PATH)
                if os.path.exists(_current_logo) and APP_CFG.get("show_logo", True):
                    st.image(_current_logo, width=90, caption="الشعار الحالي")

            _prev_company = s_company or APP_CFG.get("company_name","...")
            _prev_branch  = s_branch  or APP_CFG.get("branch_name","")
            _prev_title   = f"نموذج تقييم الأداء السنوي — {_prev_company}"
            if _prev_branch.strip():
                _prev_title += f" — {_prev_branch.strip()}"
            st.markdown(f"""<div style="background:#1F3864;color:white;padding:10px 16px;
                                border-radius:8px;text-align:center;font-size:13px;
                                font-weight:bold;margin:10px 0 4px;direction:rtl;">
                    {_prev_title}</div>
                    <p style="text-align:center;color:#888;font-size:11px;margin:2px 0 12px;">
                        ↑ معاينة ترويسة التقرير</p>""", unsafe_allow_html=True)

            if st.form_submit_button("💾 حفظ الإعدادات", use_container_width=True):
                _logo_saved_path = _current_logo
                if s_logo_file:
                    with open(LOGO_PATH, "wb") as lf:
                        lf.write(s_logo_file.getbuffer())
                    _logo_saved_path = LOGO_PATH
                    st.success("✅ تم رفع الشعار الجديد.")
                save_app_settings({
                    "company_name":   s_company,
                    "branch_name":    s_branch,
                    "employee_count": s_emp_cnt,
                    "contact_phone":  s_phone,
                    "contact_email":  s_email,
                    "show_logo":      s_show_logo,
                    "logo_path":      _logo_saved_path,
                })
                st.success("✅ تم حفظ إعدادات الشركة.")
                st.rerun()

    # ══════════════════════════════════════════════════════════════════
    # تاب 4 و 5 و 6
    # ══════════════════════════════════════════════════════════════════
    with set_tab4:
        try:
            if _EMP_KPI_OK:
                render_employees_panel()
            else:
                render_employee_management(df_emp, df_data, df_kpi, load_app_settings(), LOGO_PATH)
        except Exception as _e4:
            st.warning(f"⚠️ تعذّر تحميل إدارة الموظفين: {_e4}")
            st.info("💡 استورد البيانات من تبويب 'قاعدة البيانات' أولاً.")

    with set_tab_kpi:
        try:
            if _EMP_KPI_OK:
                render_kpis_panel()
            else:
                st.info("employees_kpis_panel.py غير موجود")
        except Exception as _ek:
            st.warning(f"⚠️ تعذّر تحميل مؤشرات الأداء: {_ek}")
            st.info("💡 استورد البيانات من تبويب 'قاعدة البيانات' أولاً.")

    with set_tab5:
        render_cv_reports(df_emp, df_data, df_kpi, load_app_settings(), LOGO_PATH)

    if _DB_PANEL_OK and set_tab6:
        with set_tab6:
            render_db_panel()
