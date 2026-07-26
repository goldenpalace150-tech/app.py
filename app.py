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

def clean_txt(raw_text):
    if not raw_text: return ""
    return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip())

@st.cache_data(ttl=300)  # Cache auth token for 5 minutes to prevent redundant login calls
def get_auth_token():
    payload = {
        "email": EMAIL,
        "password": PASSWORD,
        "company": COMPANY
    }
    try:
        response = requests.post(TOKEN_URL, json=payload, timeout=10)
        if response.status_code in:
            return response.json().get("token")
        else:
            st.error(f"فشل مصادقة API الحساب (Code: {response.status_code})")
            return None
    except Exception as e:
        st.error(f"تعذر الاتصال بـ API Token: {e}")
        return None

def load_attendance_data_from_api(selected_date_str):
    token = get_auth_token()
    if not token:
        raise Exception("Authentication Token details are missing or invalid.")
        
    headers = {
        "Authorization": f"Token {token}",
        "Content-Type": "application/json"
    }
    
    # 1. Fetch active employees
    emp_endpoint = f"{BASE_URL}/api/personnel/employees/?page_size=1000"
    emp_res = requests.get(emp_endpoint, headers=headers, timeout=15)
    all_employees = emp_res.json().get("results", []) if emp_res.status_code == 200 else []
    
    # Filter active, non-management personnel locally
    active_employees = {}
    for emp in all_employees:
        code = str(emp.get("emp_code"))
        if code not in EXCLUDED_MANAGEMENT_CODES:
            first_name = emp.get("first_name", "") or ""
            last_name = emp.get("last_name", "") or ""
            full_name = f"{first_name} {last_name}".strip()
            if not full_name:
                full_name = emp.get("nickname", "") or f"User {code}"
            active_employees[code] = clean_txt(full_name)

    # 2. Fetch specific date transactions logs
    logs_endpoint = f"{BASE_URL}/api/attendance/logs/?start_time={selected_date_str} 00:00:00&end_time={selected_date_str} 23:59:59&page_size=5000"
    logs_res = requests.get(logs_endpoint, headers=headers, timeout=15)
    raw_logs = logs_res.json().get("results", []) if logs_res.status_code == 200 else []

    # Map transaction punches per employee locally
    emp_punches = {}
    for log in raw_logs:
        code = str(log.get("emp_code"))
        if code in active_employees:
            punch_time_str = log.get("punch_time")
            try:
                p_time = datetime.strptime(punch_time_str, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                p_time = datetime.fromisoformat(punch_time_str.replace("Z", ""))
                            if code not in emp_punches:
                emp_punches[code] = []
            emp_punches[code].append(p_time)

    # Sort array punches chronologically per person
    for code in emp_punches:
        emp_punches[code].sort()

    present_staff, late_staff, full_absent_staff = [], [], []

    # Process metrics against employee configurations
    for code, name in active_employees.items():
        if code in emp_punches and emp_punches[code]:
            user_punches = emp_punches[code]
            first_punch = user_punches[0]
            punch_count = len(user_punches)
            time_in_clean = first_punch.strftime('%I:%M %p')

            # Late rule check (Entry after 09:15 AM)
            if first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15):
                late_staff.append((code, name, time_in_clean))

            # Odd/Even punch count logic (Odd count means checked in but not checked out yet)
            if punch_count % 2 != 0:
                present_staff.append((code, name, time_in_clean))
        else:
            # Employee has 0 logs registered for this date range
            full_absent_staff.append((code, name))

    # Sort results sequentially by employee code
    full_absent_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else x[0])
    present_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else x[0])
    late_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else x[0])
    
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

# Interactive local system date picker
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
