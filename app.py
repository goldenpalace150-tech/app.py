import streamlit as st
import requests
import unicodedata
from datetime import datetime
import zoneinfo

# ==========================================
# 0. RTL ARABIC TEXT CONSTANTS
# ==========================================
TEXT_CONFIG = {
    "page_title": "حضور القصر الذهبي",
    "style_align": """
        <style>
        .reportview-container .main .block-container { direction: RTL; text-align: right; }
        h1, h2, h3, h4, p, span, li, div { text-align: right !important; direction: RTL !important; line-height: 1.6 !important; }
        </style>
    """,
    "title_main": "✨ شركة القصر الذهبي ✨",
    "title_sub": "لوحة تحكم إدارة الحضور والغياب (BioTime API)",
    "lbl_date": "📅 التاريخ الحالي في سوريا: **{}**  │  ⏰ الوقت الحالي: **{}**",
    "lbl_picker": "📅 اختر التاريخ المراد عرض بياناته:",
    "btn_refresh": "🔄 تحديث البيانات الحية الآن",
    "header_late": "⏰ المتأخرون اليوم ({}) – دخول بعد 09:15 صباحاً",
    "late_row": "🔸 **{}** (كود: {}) ── وقت الدخول: {}",
    "success_no_late": "🎉 لا يوجد متأخرين في هذا التاريخ!",
    "header_absent": "❌ غائبون أو نسوا تسجيل الحضور ({})",
    "absent_row": "🔹 **{}** (كود: {})",
    "success_no_absent": "🎉 لا يوجد غيابات في هذا التاريخ!",
    "header_present": "🟢 الموظفون المتواجدون حالياً في العمل ({})",
    "present_row": "🔸 **{}** (كود: {}) ── وقت الدخول: {}",
    "info_no_present": "لا يوجد موظفين منتظمين متواجدين حالياً.",
    "err_api": "خطأ في الاتصال بواجهة BioTime السحابية: {}"
}

st.set_page_config(page_title=TEXT_CONFIG["page_title"], page_icon="📊", layout="wide")
st.markdown(TEXT_CONFIG["style_align"], unsafe_allow_html=True)

# Configurations
EXCLUDED_MANAGEMENT_CODES = ("40", "10", "20")
SYRIA_TZ = zoneinfo.ZoneInfo("Asia/Damascus")

# Retrieve API secrets dynamically from Streamlit Secrets Environment
BASE_URL = st.secrets["biotime"]["base_url"].rstrip('/')
TOKEN_URL = st.secrets["biotime"]["token_url"]
EMAIL = st.secrets["biotime"]["email"]
PASSWORD = st.secrets["biotime"]["password"]
COMPANY = st.secrets["biotime"]["company"]

# Initialize a debug logger list in session state
if "debug_logs" not in st.session_state:
    st.session_state["debug_logs"] = []

def log_debug(message):
    st.session_state["debug_logs"].append(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

def clean_txt(raw_text):
    if not raw_text: return ""
    return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip())
@st.cache_data(ttl=300)  # Cache auth token for 5 minutes
def get_auth_token():
    payload = {
        "email": EMAIL,
        "password": PASSWORD,
        "company": COMPANY
    }
    try:
        response = requests.post(TOKEN_URL, json=payload, timeout=10)
        if response.status_code == 200 or response.status_code == 201:
            token = response.json().get("token")
            log_debug(f"Auth Success! Token truncated: {token[:10]}...")
            return token
        else:
            st.error(f"فشل مصادقة API الحساب (Code: {response.status_code})")
            log_debug(f"Auth Failed Code {response.status_code}: {response.text}")
            return None
    except Exception as e:
        st.error(f"تعذر الاتصال بـ API Token: {e}")
        log_debug(f"Auth Exception: {str(e)}")
        return None

def load_attendance_data_from_api(selected_date_str):
    token = get_auth_token()
    if not token:
        raise Exception("Authentication Token details are missing or invalid.")
        
    headers = {
        "Authorization": f"Token {token}",
        "Content-Type": "application/json"
    }
    
    st.session_state["debug_logs"] = [] # Reset logs on refresh
    
    # Fetch Employees (With fallback paths)
    emp_endpoints = [
        f"{BASE_URL}/api/personnel/employees/?page_size=1000",
        f"{BASE_URL}/api/employees/?page_size=1000"
    ]
    
    all_employees = []
    for url in emp_endpoints:
        log_debug(f"Trying Employee Endpoint: {url}")
        try:
            emp_res = requests.get(url, headers=headers, timeout=15)
            log_debug(f"Employee Endpoint Response Status: {emp_res.status_code}")
            if emp_res.status_code == 200:
                res_json = emp_res.json()
                all_employees = res_json.get("results", res_json if isinstance(res_json, list) else [])
                log_debug(f"Found {len(all_employees)} raw employee profiles.")
                if all_employees:
                    break
        except Exception as e:
            log_debug(f"Employee Endpoint Error: {str(e)}")

    active_employees = {}
    for emp in all_employees:
        code = str(emp.get("emp_code", emp.get("pin", "")))
        if code and code not in EXCLUDED_MANAGEMENT_CODES:
            first_name = emp.get("first_name", "") or ""
            last_name = emp.get("last_name", "") or ""
            full_name = f"{first_name} {last_name}".strip()
            if not full_name:
                full_name = emp.get("nickname", "") or emp.get("name", "") or f"User {code}"
            active_employees[code] = clean_txt(full_name)

    log_debug(f"Filtered to {len(active_employees)} active processing tracking codes.")

    # Fetch Logs (With fallback paths)
    logs_endpoints = [
        f"{BASE_URL}/api/attendance/logs/?start_time={selected_date_str} 00:00:00&end_time={selected_date_str} 23:59:59&page_size=5000",
        f"{BASE_URL}/api/iclock/transactions/?start_time={selected_date_str} 00:00:00&end_time={selected_date_str} 23:59:59&page_size=5000",
        f"{BASE_URL}/api/transactions/?punch_date={selected_date_str}&page_size=5000"
    ]
    
    raw_logs = []
    for url in logs_endpoints:
        log_debug(f"Trying Logs Endpoint: {url}")
        try:
            logs_res = requests.get(url, headers=headers, timeout=15)
            log_debug(f"Logs Endpoint Response Status: {logs_res.status_code}")
            if logs_res.status_code == 200:
                res_json = logs_res.json()
                raw_logs = res_json.get("results", res_json if isinstance(res_json, list) else [])
                log_debug(f"Found {len(raw_logs)} raw punch items.")
                if raw_logs:
                    break
        except Exception as e:
            log_debug(f"Logs Endpoint Error: {str(e)}")

    emp_punches = {}
    for log in raw_logs:
        code = str(log.get("emp_code", log.get("pin", log.get("emp_id", ""))))
        if code in active_employees:
            punch_time_str = log.get("punch_time", log.get("punch_date", ""))
            if not punch_time_str:
                continue
            try:
                p_time = datetime.strptime(punch_time_str[:19], "%Y-%m-%d %H:%M:%S")
            except ValueError:
                try:
                    p_time = datetime.fromisoformat(punch_time_str.replace("Z", ""))
                except Exception:
                    continue
                
            if code not in emp_punches:
                emp_punches[code] = []
            emp_punches[code].append(p_time)

    for code in emp_punches:
        emp_punches[code].sort()

    present_staff, late_staff, full_absent_staff = [], [], []

    for code, name in active_employees.items():
        if code in emp_punches and emp_punches[code]:
            user_punches = emp_punches[code]
            first_punch = user_punches
            punch_count = len(user_punches)
            time_in_clean = first_punch.strftime('%I:%M %p')

            if first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15):
                late_staff.append((code, name, time_in_clean))

            if punch_count % 2 != 0:
                present_staff.append((code, name, time_in_clean))
        else:
            full_absent_staff.append((code, name))

    full_absent_staff.sort(key=lambda x: int(x) if x.isdigit() else x)
    present_staff.sort(key=lambda x: int(x) if x.isdigit() else x)
    late_staff.sort(key=lambda x: int(x) if x.isdigit() else x)
    
    return present_staff, late_staff, full_absent_staff
# ==========================================
# 3. INTERFACE RENDERING
# ==========================================
now_syria = datetime.now(SYRIA_TZ)
current_today = now_syria.date()
time_str = now_syria.strftime('%I:%M:%S %p')

st.title(TEXT_CONFIG["title_main"])
st.subheader(TEXT_CONFIG["title_sub"])
st.markdown(TEXT_CONFIG["lbl_date"].format(current_today.strftime('%Y-%m-%d'), time_str))

selected_date = st.date_input(TEXT_CONFIG["lbl_picker"], value=current_today)
selected_date_str = selected_date.strftime('%Y-%m-%d')

if st.button(TEXT_CONFIG["btn_refresh"]):
    st.cache_data.clear()  # Clear API query cache records completely
    st.rerun()

try:
    present_staff, late_staff, full_absent_staff = load_attendance_data_from_api(selected_date_str)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(TEXT_CONFIG["header_late"].format(len(late_staff)))
        if late_staff:
            for emp_code, name, time_in in late_staff:
                st.markdown(TEXT_CONFIG["late_row"].format(name, emp_code, time_in))
        else:
            st.success(TEXT_CONFIG["success_no_late"])
            
    with col2:
        st.markdown(TEXT_CONFIG["header_present"].format(len(present_staff)))
        if present_staff:
            for emp_code, name, time_in in present_staff:
                st.markdown(TEXT_CONFIG["present_row"].format(name, emp_code, time_in))
        else:
            st.info(TEXT_CONFIG["info_no_present"])
            
    with col3:
        st.markdown(TEXT_CONFIG["header_absent"].format(len(full_absent_staff)))
        if full_absent_staff:
            for emp_code, name in full_absent_staff:
                st.markdown(TEXT_CONFIG["absent_row"].format(name, emp_code))
        else:
            st.success(TEXT_CONFIG["success_no_absent"])
            
except Exception as e:
    st.error(TEXT_CONFIG["err_api"].format(str(e)))

# ==========================================
# 4. LIVE DEVELOPER DEBUG PANEL (BOTTOM)
# ==========================================
st.write("---")
with st.expander("🛠️ معلومات التصحيح المباشرة / Live API Debug Logger"):
    st.write("If values are zero, check which endpoints below returned 0 rows or error codes:")
    for log in st.session_state.get("debug_logs", []):
        st.text(log)
