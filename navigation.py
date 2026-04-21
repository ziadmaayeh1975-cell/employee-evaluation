import os
from datetime import date
import streamlit as st
from constants import LOGO_PATH

def render_sidebar(users):
    disp = users.get(st.session_state.username, {}).get("display", st.session_state.username)
    with st.sidebar:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=130)
        st.markdown(f"""
        <div style="background:#1E3A8A;border-radius:10px;padding:12px 14px;margin:8px 0;">
            <p style="color:#BDD7EE;font-size:12px;margin:0;">مرحباً،</p>
            <p style="color:white;font-size:15px;font-weight:bold;margin:2px 0;">👤 {disp}</p>
        </div>""", unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("<p style='color:#64748B;font-size:12px;font-weight:bold;margin:0 0 6px;'>📌 القائمة الرئيسية</p>", unsafe_allow_html=True)
        pages = {
            "📝 إدخال التقييم": "entry",
            "👁️ عرض وتعديل وحذف التقييم": "manage",
        }
        report_pages = {
            "📄 تقرير موظف": "rep_emp",
            "📊 تقرير مفصّل بالأقسام": "rep_dept",
            "🏆 الملخص السنوي": "rep_year",
        }
        admin_pages = {"⚙️ الإعدادات": "settings"} if st.session_state.role in ("admin", "super_admin") else {}
        if "page" not in st.session_state:
            st.session_state.page = "entry"
        def nav_btn(label, key):
            if st.button(label, key=f"nav_{key}", use_container_width=True):
                st.session_state.page = key
                st.rerun()
        for label, key in pages.items():
            nav_btn(label, key)
        st.markdown("<p style='color:#64748B;font-size:12px;font-weight:bold;margin:10px 0 6px;'>📊 التقارير</p>", unsafe_allow_html=True)
        for label, key in report_pages.items():
            nav_btn(label, key)
        if admin_pages:
            st.markdown("<p style='color:#64748B;font-size:12px;font-weight:bold;margin:10px 0 6px;'>🔧 الإدارة</p>", unsafe_allow_html=True)
            for label, key in admin_pages.items():
                nav_btn(label, key)
        st.markdown("---")
        if st.button("🚪 تسجيل الخروج", use_container_width=True, key="logout_btn"):
            st.session_state.logged_in = False
            st.session_state.username = ""
            st.session_state.role = ""
            st.rerun()
    return st.session_state.page

def render_page_header(page):
    page_titles = {
        "entry": ("📝 إدخال التقييم الشهري", "#1E3A8A"),
        "manage": ("👁️ عرض وتعديل وحذف التقييم", "#7C3AED"),
        "rep_emp": ("📄 تقرير الموظف", "#0F766E"),
        "rep_dept": ("📊 التقرير المفصّل بالأقسام", "#0369A1"),
        "rep_year": ("🏆 الملخص السنوي", "#B45309"),
        "settings": ("⚙️ إعدادات النظام", "#DC2626"),
    }
    pg_title, pg_color = page_titles.get(page, ("النظام", "#1E3A8A"))
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,{pg_color}15,{pg_color}05);
                border-right:5px solid {pg_color};border-radius:0 10px 10px 0;
                padding:14px 20px;margin-bottom:20px;">
        <h2 style="margin:0;color:{pg_color};font-size:1.4rem;">{pg_title}</h2>
        <p style="margin:2px 0 0;color:#64748B;font-size:12px;">
            نظام تقييم الأداء – مجموعة شركات فنون &nbsp;|&nbsp; {date.today().strftime("%d/%m/%Y")}
        </p>
    </div>""", unsafe_allow_html=True)

