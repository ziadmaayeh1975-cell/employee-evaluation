import hashlib
import json
import os
from datetime import datetime
import streamlit as st
from constants import LOGO_PATH, SETTINGS_FILE, TRIAL_FILE, USERS_FILE

def hash_pw(pw):
    return hashlib.sha256(pw.encode()).hexdigest()

def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"admin": {"password": hash_pw("admin123"), "role": "admin", "display": "المدير"}}

def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

def check_login(username, password, users):
    u = users.get(username.strip())
    return u and u["password"] == hash_pw(password)

def load_app_settings():
    defaults = {
        "company_name": "مجموعة شركات فنون",
        "employee_count": "",
        "contact_phone": "",
        "contact_email": "",
        "logo_path": "logo.png",
        "show_logo": True,
    }
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            saved = json.load(f)
            defaults.update(saved)
    return defaults

def save_app_settings(settings):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)

def load_trial_users():
    if os.path.exists(TRIAL_FILE):
        with open(TRIAL_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_trial_users(trial):
    with open(TRIAL_FILE, "w", encoding="utf-8") as f:
        json.dump(trial, f, ensure_ascii=False, indent=2)

def check_trial_access(username):
    trial = load_trial_users()
    if username not in trial:
        return True, ""
    t = trial[username]
    try:
        start = datetime.strptime(t["start"], "%Y-%m-%d %H:%M")
        end = datetime.strptime(t["end"], "%Y-%m-%d %H:%M")
        now = datetime.now()
        if now < start:
            return False, f"⏳ فترة التجربة لم تبدأ بعد. تبدأ في {t['start']}"
        if now > end:
            return False, "expired"
        return True, ""
    except:
        return True, ""

def ensure_session_state():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
    if "username" not in st.session_state:
        st.session_state.username = ""
    if "role" not in st.session_state:
        st.session_state.role = ""
    if "is_trial" not in st.session_state:
        st.session_state.is_trial = False

def render_login(users):
    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=120)
        st.markdown("<h2 style='text-align:center;color:#1E3A8A;'>نظام تقييم الأداء</h2>", unsafe_allow_html=True)
        st.markdown("<h4 style='text-align:center;color:#475569;'>مجموعة شركات فنون</h4>", unsafe_allow_html=True)
        st.write("")
        with st.form("login_form"):
            uname = st.text_input("👤 اسم المستخدم", placeholder="أدخل اسم المستخدم")
            passw = st.text_input("🔒 كلمة المرور", type="password", placeholder="أدخل كلمة المرور")
            if st.form_submit_button("🚀 دخول", use_container_width=True):
                if check_login(uname, passw, users):
                    allowed, msg = check_trial_access(uname.strip())
                    if not allowed:
                        if msg == "expired":
                            st.markdown("""
                            <div style="text-align:center;padding:40px 20px;
                                        background:#FFF7ED;border-radius:16px;
                                        border:2px solid #F59E0B;">
                                <div style="font-size:48px;">⏰</div>
                                <h2 style="color:#B45309;">انتهت فترة التجربة</h2>
                                <p style="color:#92400E;font-size:16px;">
                                    نشكرك على استخدام برنامج التقييم السنوي للموظفين
                                </p>
                                <p style="color:#64748B;font-size:13px;">
                                    للحصول على النسخة الكاملة، يرجى التواصل مع الإدارة.
                                </p>
                            </div>""", unsafe_allow_html=True)
                        else:
                            st.warning(msg)
                    else:
                        st.session_state.logged_in = True
                        st.session_state.username = uname.strip()
                        st.session_state.role = users[uname.strip()]["role"]
                        st.session_state.is_trial = uname.strip() in load_trial_users()
                        st.rerun()
                else:
                    st.error("❌ اسم المستخدم أو كلمة المرور غير صحيحة.")

def stop_if_trial_expired():
    if st.session_state.get("is_trial"):
        ok, msg = check_trial_access(st.session_state.username)
        if not ok:
            st.session_state.logged_in = False
            st.session_state.is_trial = False
            app_settings = load_app_settings()
            st.markdown(f"""
            <div style="display:flex;justify-content:center;align-items:center;
                        min-height:60vh;">
            <div style="text-align:center;padding:50px 40px;background:white;
                        border-radius:20px;box-shadow:0 4px 24px rgba(0,0,0,0.12);
                        max-width:480px;width:100%;">
                <div style="font-size:56px;margin-bottom:12px;">⏰</div>
                <h2 style="color:#1E3A8A;margin-bottom:8px;">انتهت فترة التجربة</h2>
                <div style="width:60px;height:3px;background:#ED7D31;
                            margin:0 auto 16px;border-radius:2px;"></div>
                <p style="color:#475569;font-size:15px;line-height:1.7;">
                    نشكرك على استخدام<br>
                    <b>برنامج التقييم السنوي للموظفين</b><br>
                    {app_settings.get('company_name','')}
                </p>
                <p style="color:#94A3B8;font-size:12px;margin-top:16px;">
                    للحصول على النسخة الكاملة، يرجى التواصل مع الإدارة.
                </p>
            </div></div>""", unsafe_allow_html=True)
            st.stop()



def get_current_reviewer():
    """يُعيد اسم المقيم المرتبط بالمستخدم الحالي، أو فارغ إذا كان admin"""
    import streamlit as st
    users = load_users()
    uname = st.session_state.get("username", "")
    return users.get(uname, {}).get("reviewer", "")


def get_current_role():
    """يُعيد دور المستخدم الحالي"""
    import streamlit as st
    return st.session_state.get("role", "user")

def is_super_admin():
    """هل المستخدم الحالي أدمن رئيسي؟"""
    import streamlit as st
    return st.session_state.get("role") == "super_admin"

def is_any_admin():
    """هل المستخدم أدمن عادي أو رئيسي؟"""
    import streamlit as st
    return st.session_state.get("role") in ("admin", "super_admin")
