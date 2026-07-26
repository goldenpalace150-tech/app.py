import streamlit as st
import requests
import unicodedata
from datetime import datetime
import zoneinfo

# ==========================================
# 0. RTL ARABIC TEXT & VISUAL CONFIG
# ==========================================
TEXT_CONFIG = {
    "page_title": "حضور القصر الذهبي",
    "title_main": "✨ شركة القصر الذهبي ✨",
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

# Custom CSS styling to recreate the look of the BioTime dashboard cards
st.markdown("""
    <style>
    .reportview-container .main .block-container { direction: RTL; text-align: right; }
    h1, h2, h3, h4, p, span, li, div { text-align: right !important; direction: RTL !important; line-height: 1.6 !important; }
    
    /* Dashboard Card Styling */
    .metric-card {
        background-color: #ffffff;
        border-radius: 8px;
        padding: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08);
        border: 1px solid #eef2f5;
        display: flex;
        align-items: center;
        margin-bottom: 10px;
    }
    .metric-icon {
        width: 45px;
        height: 45px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 22px;
        margin-left: 15px;
    }
    .icon-total { background-color: #e8f5e9; color: #4caf50; }
    .icon-present { background-color: #e8f5e9; color: #4caf50; }
    .icon-absent { background-color: #ffebee; color: #f44336; }
    .icon-late { background-color: #fff8e1; color: #ffc107; }
    
    .metric-info {
        display: flex;
        flex-direction: column;
    }
    .metric-title {
        font-size: 14px;
        color: #64748b;
        font-weight: 500;
    }
    .metric-value {
        font-size: 24px;
        font-weight: bold;
        color: #1e293b;
    }
    </style>
""", unsafe_allow_html=True)

# Configurations
EXCLUDED_MANAGEMENT_CODES = ("40", "10", "20")
SYRIA_TZ = zoneinfo.ZoneInfo("Asia/Damascus")

# Retrieve API secrets dynamically from Streamlit Secrets Environment
BASE_URL = st.secrets["biotime"]["base_url"].rstrip('/')
TOKEN_URL = st.secrets["biotime"]["token_url"]
EMAIL = st.secrets["biotime"]["email"]
PASSWORD = st.secrets["biotime"]["password"]
COMPANY = st.secrets["biotime"]["company"]

if "debug_logs" not in st.session_state:
    st.session_state["debug_logs"] = []

# Keep track of which card tile is selected clicked
if "active_view" not in st.session_state:
    st.session_state["active_view"] = "present"

def log_debug(message):
    st.session_state["debug_logs"].append(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

def clean_txt(raw_text):
    if not raw_text: return ""
    return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip())
@st.cache_data(ttl=300)
def get_auth_token():
    payload = {"email": EMAIL, "password": PASSWORD, "company": COMPANY}
    try:
        response = requests.post(TOKEN_URL, json=payload, timeout=10)
        if response.status_code == 200 or response.status_code == 201:
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
        
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    st.session_state["debug_logs"] = []
    
    # 1. Fetch Employee List
    emp_url = f"{BASE_URL}/personnel/api/employees/?page_size=1000"
    all_employees = []
    try:
        emp_res = requests.get(emp_url, headers=headers, timeout=15)
        if emp_res.status_code == 200:
            all_employees = emp_res.json().get("data", [])
    except Exception as e:
        log_debug(f"Employee Request Error: {str(e)}")

    active_employees = {}
    for emp in all_employees:
        code = str(emp.get("emp_code", ""))
        if code and code not in EXCLUDED_MANAGEMENT_CODES:
            first_name = emp.get("first_name", "") or ""
            last_name = emp.get("last_name", "") or ""
            full_name = f"{first_name} {last_name}".strip()
            active_employees[code] = clean_txt(full_name if full_name else f"User {code}")

    # 2. Fetch Transaction Logs
    logs_url = f"{BASE_URL}/iclock/api/transactions/?start_time={selected_date_str} 00:00:00&end_time={selected_date_str} 23:59:59&page_size=5000"
    raw_logs = []
    try:
        logs_res = requests.get(logs_url, headers=headers, timeout=15)
        if logs_res.status_code == 200:
            raw_logs = logs_res.json().get("data", [])
    except Exception as e:
        log_debug(f"Logs Request Error: {str(e)}")

    emp_punches = {}
    for log in raw_logs:
        code = str(log.get("emp_code", ""))
        if code in active_employees:
            punch_time_str = log.get("punch_time", "")
            if punch_time_str:
                try:
                    p_time = datetime.strptime(punch_time_str[:19], "%Y-%m-%d %H:%M:%S")
                    if code not in emp_punches:
                        emp_punches[code] = []
                    emp_punches[code].append(p_time)
                except Exception:
                    continue

    for code in emp_punches:
        emp_punches[code].sort()

    present_staff, late_staff, full_absent_staff = [], [], []

    for code, name in active_employees.items():
        if code in emp_punches and emp_punches[code]:
            user_punches = emp_punches[code]
            first_punch = user_punches[0]
            punch_count = len(user_punches)
            time_in_clean = first_punch.strftime('%I:%M %p')

            if first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15):
                late_staff.append((code, name, time_in_clean))

            if punch_count % 2 != 0:
                present_staff.append((code, name, time_in_clean))
        else:
            full_absent_staff.append((code, name))

    full_absent_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else x[0])
    present_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else x[0])
    late_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else x[0])
    
    return len(active_employees), present_staff, late_staff, full_absent_staff
# ==========================================
# 3. INTERFACE RENDERING
# ==========================================
now_syria = datetime.now(SYRIA_TZ)
current_today = now_syria.date()
time_str = now_syria.strftime('%I:%M:%S %p')

st.title(TEXT_CONFIG["title_main"])
st.markdown(TEXT_CONFIG["lbl_date"].format(current_today.strftime('%Y-%m-%d'), time_str))

# Interactive controls layout
c_date, c_ref = st.columns([4, 1])
with c_date:
    selected_date = st.date_input(TEXT_CONFIG["lbl_picker"], value=current_today)
    selected_date_str = selected_date.strftime('%Y-%m-%d')
with c_ref:
    st.write("<br>", unsafe_allow_html=True)
    if st.button(TEXT_CONFIG["btn_refresh"], use_container_width=True):
        st.cache_data.clear()
        st.rerun()

try:
    total_emp, present_staff, late_staff, full_absent_staff = load_attendance_data_from_api(selected_date_str)
    
    # ---------------------------------------------
    # RENDER CLICKABLE STATISTIC CARDS (LIKE BIOTIME)
    # ---------------------------------------------
    st.write("### 📊 إحصائيات عامة / OVERALL STATISTICS")
    card_col1, card_col2, card_col3, card_col4 = st.columns(4)
    
    with card_col1:
        st.markdown(f'''
            <div class="metric-card">
                <div class="metric-icon icon-total">👥</div>
                <div class="metric-info">
                    <span class="metric-title">Employees</span>
                    <span class="metric-value">{total_emp}</span>
                </div>
            </div>
        ''', unsafe_allow_html=True)
        st.button("📄 عرض جميع الموظفين", key="btn_view_total", on_click=lambda: st.session_state.update({"active_view": "total"}), use_container_width=True)

    with card_col2:
        st.markdown(f'''
            <div class="metric-card">
                <div class="metric-icon icon-present">👤</div>
                <div class="metric-info">
                    <span class="metric-title">Present</span>
                    <span class="metric-value">{len(present_staff)}</span>
                </div>
            </div>
        ''', unsafe_allow_html=True)
        st.button("🟢 عرض المتواجدين حالياً", key="btn_view_present", on_click=lambda: st.session_state.update({"active_view": "present"}), use_container_width=True)

    with card_col3:
        st.markdown(f'''
            <div class="metric-card">
                <div class="metric-icon icon-absent">📅</div>
                <div class="metric-info">
                    <span class="metric-title">Absent</span>
                    <span class="metric-value">{len(full_absent_staff)}</span>
                </div>
            </div>
        ''', unsafe_allow_html=True)
        st.button("❌ عرض الغائبين اليوم", key="btn_view_absent", on_click=lambda: st.session_state.update({"active_view": "absent"}), use_container_width=True)

    with card_col4:
        st.markdown(f'''
            <div class="metric-card">
                <div class="metric-icon icon-late">⏳</div>
                <div class="metric-info">
                    <span class="metric-title">Late Arrival</span>
                    <span class="metric-value">{len(late_staff)}</span>
                </div>
            </div>
        ''', unsafe_allow_html=True)
        st.button("⏰ عرض المتأخرين", key="btn_view_late", on_click=lambda: st.session_state.update({"active_view": "late"}), use_container_width=True)

    # ---------------------------------------------
    # DYNAMIC INTERACTIVE EMPLOYEE VIEW LIST BOX
    # ---------------------------------------------
    st.write("---")
    current_view = st.session_state["active_view"]
    
    if current_view == "present":
        st.markdown(TEXT_CONFIG["header_present"].format(len(present_staff)))
        if present_staff:
            for emp_code, name, time_in in present_staff:
                st.markdown(TEXT_CONFIG["present_row"].format(name, emp_code, time_in))
        else:
            st.info(TEXT_CONFIG["info_no_present"])
            
    elif current_view == "absent":
        st.markdown(TEXT_CONFIG["header_absent"].format(len(full_absent_staff)))
        if full_absent_staff:
            for emp_code, name in full_absent_staff:
                st.markdown(TEXT_CONFIG["absent_row"].format(name, emp_code))
        else:
            st.success(TEXT_CONFIG["success_no_absent"])
            
    elif current_view == "late":
        st.markdown(TEXT_CONFIG["header_late"].format(len(late_staff)))
        if late_staff:
            for emp_code, name, time_in in late_staff:
                st.markdown(TEXT_CONFIG["late_row"].format(name, emp_code, time_in))
        else:
            st.success(TEXT_CONFIG["success_no_late"])
            
    elif current_view == "total":
        st.markdown("### 📋 قائمة الموظفين الكاملة النشطة للتتبع")
        token = get_auth_token()
        headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
        emp_res = requests.get(f"{BASE_URL}/personnel/api/employees/?page_size=1000", headers=headers).json().get("data", [])
        for emp in emp_res:
            code = str(emp.get("emp_code", ""))
            if code not in EXCLUDED_MANAGEMENT_CODES:
                st.markdown(f"🔹 **{emp.get('first_name','')} {emp.get('last_name','') or ''}** (كود الموظف: {code})")
            
except Exception as e:
    st.error(TEXT_CONFIG["err_api"].format(str(e)))

# ==========================================
# 4. LIVE DEVELOPER DEBUG PANEL (BOTTOM)
# ==========================================
st.write("---")
with st.expander("🛠️ معلومات التصحيح المباشرة / Live API Debug Logger"):
    for log in st.session_state.get("debug_logs", []):
        st.text(log)
