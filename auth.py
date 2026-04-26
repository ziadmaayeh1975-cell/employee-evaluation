import hashlib
import json
import os
from datetime import datetime
import streamlit as st
from constants import LOGO_PATH, SETTINGS_FILE, TRIAL_FILE, USERS_FILE

def hash_pw(pw):
    return hashlib.sha256(pw.encode()).hexdigest()

# ── اكتشاف Supabase ──────────────────────────────────────────────
def _use_supabase():
    try:
        url = st.secrets.get("SUPABASE_URL","") or os.environ.get("SUPABASE_URL","")
        key = st.secrets.get("SUPABASE_KEY","") or os.environ.get("SUPABASE_KEY","")
        return bool(url and key)
    except:
        return False

def _get_supabase():
    from supabase import create_client
    url = st.secrets.get("SUPABASE_URL","") or os.environ.get("SUPABASE_URL","")
    key = st.secrets.get("SUPABASE_KEY","") or os.environ.get("SUPABASE_KEY","")
    return create_client(url, key)

# ════════════════════════════════════════════════════════════════
# إدارة المستخدمين
# ════════════════════════════════════════════════════════════════
def load_users():
    if _use_supabase():
        try:
            sb = _get_supabase()
            rows = sb.table("users").select("*").execute().data
            return {r["username"]: {
                "password": r["password"],
                "role":     r.get("role","user"),
                "display":  r.get("display", r["username"]),
                "reviewer": r.get("reviewer",""),
            } for r in rows}
        except Exception as e:
            st.warning(f"⚠️ تعذّر تحميل المستخدمين من Supabase: {e}")
    # JSON محلي
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"admin": {"password": hash_pw("admin123"), "role": "super_admin", "display": "المدير", "reviewer": ""}}


def save_users(users):
    if _use_supabase():
        try:
            sb = _get_supabase()
            # احذف الكل وأعد الإدراج
            sb.table("users").delete().neq("username","__none__").execute()
            for uname, udata in users.items():
                sb.table("users").insert({
                    "username": uname,
                    "password": udata["password"],
                    "role":     udata.get("role","user"),
                    "display":  udata.get("display", uname),
                    "reviewer": udata.get("reviewer",""),
                }).execute()
            return
        except Exception as e:
            st.warning(f"⚠️ تعذّر حفظ المستخدمين في Supabase: {e}")
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)


def add_user(username, password, role, display, reviewer=""):
    """إضافة مستخدم جديد."""
    if _use_supabase():
        try:
            sb = _get_supabase()
            sb.table("users").insert({
                "username": username,
                "password": hash_pw(password),
                "role":     role,
                "display":  display,
                "reviewer": reviewer,
            }).execute()
            return True, None
        except Exception as e:
            return False, str(e)
    users = load_users()
    if username in users:
        return False, "المستخدم موجود مسبقاً"
    users[username] = {"password": hash_pw(password), "role": role,
                       "display": display, "reviewer": reviewer}
    save_users(users)
    return True, None


def update_user(username, data):
    """تعديل بيانات مستخدم."""
    if _use_supabase():
        try:
            sb = _get_supabase()
            sb.table("users").update(data).eq("username", username).execute()
            return True, None
        except Exception as e:
            return False, str(e)
    users = load_users()
    if username not in users:
        return False, "المستخدم غير موجود"
    users[username].update(data)
    save_users(users)
    return True, None


def delete_user(username):
    """حذف مستخدم."""
    if _use_supabase():
        try:
            sb = _get_supabase()
            sb.table("users").delete().eq("username", username).execute()
            return True, None
        except Exception as e:
            return False, str(e)
    users = load_users()
    if username in users:
        del users[username]
        save_users(users)
    return True, None


def check_login(username, password, users):
    u = users.get(username.strip())
    return u and u["password"] == hash_pw(password)


# ════════════════════════════════════════════════════════════════
# إعدادات التطبيق
# ════════════════════════════════════════════════════════════════
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


# ════════════════════════════════════════════════════════════════
# المستخدمون التجريبيون
# ════════════════════════════════════════════════════════════════
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
        end   = datetime.strptime(t["end"],   "%Y-%m-%d %H:%M")
        now   = datetime.now()
        if now < start: return False, f"⏳ فترة التجربة لم تبدأ بعد. تبدأ في {t['start']}"
        if now > end:   return False, "expired"
        return True, ""
    except:
        return True, ""


# ════════════════════════════════════════════════════════════════
# Session State
# ════════════════════════════════════════════════════════════════
def ensure_session_state():
    for key, val in [("logged_in",False),("username",""),("role",""),("is_trial",False)]:
        if key not in st.session_state:
            st.session_state[key] = val


def render_login(users):
    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=120)
        st.markdown("<h2 style='text-align:center;color:#1E3A8A;'>برنامج تقييم الموظفين</h2>", unsafe_allow_html=True)
        st.write("")
        with st.form("login_form"):
            uname = st.text_input("👤 اسم المستخدم", placeholder="أدخل اسم المستخدم")
            passw = st.text_input("🔒 كلمة المرور", type="password", placeholder="أدخل كلمة المرور")
            if st.form_submit_button("🚀 دخول", use_container_width=True):
                if check_login(uname, passw, users):
                    allowed, msg = check_trial_access(uname.strip())
                    if not allowed:
                        if msg == "expired":
                            st.error("⏰ انتهت فترة التجربة.")
                        else:
                            st.warning(msg)
                    else:
                        st.session_state.logged_in = True
                        st.session_state.username  = uname.strip()
                        st.session_state.role      = users[uname.strip()]["role"]
                        st.session_state.is_trial  = uname.strip() in load_trial_users()
                        st.rerun()
                else:
                    st.error("❌ اسم المستخدم أو كلمة المرور غير صحيحة.")


def stop_if_trial_expired():
    if st.session_state.get("is_trial"):
        ok, msg = check_trial_access(st.session_state.username)
        if not ok:
            st.session_state.logged_in = False
            st.session_state.is_trial  = False
            st.error("⏰ انتهت فترة التجربة.")
            st.stop()


def get_current_reviewer():
    users = load_users()
    uname = st.session_state.get("username","")
    return users.get(uname, {}).get("reviewer","")

def get_current_role():
    return st.session_state.get("role","user")

def is_super_admin():
    return st.session_state.get("role") == "super_admin"

def is_any_admin():
    return st.session_state.get("role") in ("admin","super_admin")
