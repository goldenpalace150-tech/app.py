import streamlit as st
import requests
import unicodedata
import pandas as pd
import io
from datetime import datetime
import zoneinfo

# ==========================================
# 0. RTL ARABIC TEXT & VISUAL CONFIG
# ==========================================
TEXT_CONFIG = {
    "page_title": "Golden Palace Attendance Dashboard",
    "title_main": "✨ Golden Palace Co. ✨",
    "lbl_date": "📅 Syria Current Date: **{}**  │  ⏰ Current Time: **{}**",
    "btn_refresh": "🔄 Refresh Live Data Now",
    "btn_download_excel": "📥 Download Daily Attendance Report (Excel Matrix)",
    
    # Section Accordion Headers
    "header_late": "⏰ Late Staff Today ({})",
    "header_absent": "❌ Full Absence Today ({})",
    "header_present": "🟢 Present / Active On-Premises ({})",
    "header_checkout": "🏁 Checked-Out / Workday Finished ({})",
    "header_all": "👥 All Active Registered Staff ({})",
    
    # Inline Rows Formatting
    "late_row": "🔸 **{}** (ID: {}) ── Clock In: {}",
    "absent_row": "🔹 **{}** (ID: {})",
    "present_row": "🔸 **{}** (ID: {}) ── Clock In: {}",
    "checkout_row": "✅ **{}** (ID: {}) ── Clock Out: {}",
    "all_row": "👤 **{}** (ID: {})",
    
    "err_api": "Cloud BioTime API Connectivity Failure: {}"
}

st.set_page_config(page_title=TEXT_CONFIG["page_title"], page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; }
    .reportview-container .main .block-container { direction: RTL; text-align: right; }
    h1, h2, h3, h4, p, span, li, div { text-align: right !important; direction: RTL !important; line-height: 1.6 !important; }
    
    /* Layout styling matching mobile display viewports */
    div.stButton > button {
        width: 100% !important;
        background-color: #ffffff !important;
        color: #1e293b !important;
        border: 1px solid #eef2f5 !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.06) !important;
        border-radius: 8px !important;
        padding: 15px !important;
        text-align: right !important;
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        margin-bottom: 2px !important;
    }
    div.stButton > button:hover {
        border-color: #cbd5e1 !important;
        background-color: #f8fafc !important;
    }
    
    .list-wrapper-box {
        background-color: #f8fafc;
        border: 1px dashed #cbd5e1;
        border-radius: 6px;
        padding: 15px;
        margin-top: 5px;
        margin-bottom: 15px;
        direction: rtl;
    }
    </style>
""", unsafe_allow_html=True)

# Define Excluded Administrative Management Codes
EXCLUDED_MANAGEMENT_CODES = ("40", "10", "20")

# HARD EXCLUSION: Add IDs of resigned or disabled employees here to strip them completely out.
EXCLUDED_RESIGNED_CODES = ("28", "34") 

SYRIA_TZ = zoneinfo.ZoneInfo("Asia/Damascus")

BASE_URL = st.secrets["biotime"]["base_url"].rstrip('/')
TOKEN_URL = st.secrets["biotime"]["token_url"]
EMAIL = st.secrets["biotime"]["email"]
PASSWORD = st.secrets["biotime"]["password"]
COMPANY = st.secrets["biotime"]["company"]

if "debug_logs" not in st.session_state:
    st.session_state["debug_logs"] = []

if "selected_view" not in st.session_state:
    st.session_state["selected_view"] = "present"

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
        return None
    except Exception:
        return None

def load_attendance_data_from_api(selected_date_str):
    token = get_auth_token()
    if not token:
        raise Exception("Authentication Token is missing or invalid.")
        
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    st.session_state["debug_logs"] = []
    
    # 1. Fetch active company staff profile definitions
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
        # Strict checking filters drop management and specified inactive profiles instantly
        if code and code not in EXCLUDED_MANAGEMENT_CODES and code not in EXCLUDED_RESIGNED_CODES:
            first_name = emp.get("first_name", "") or ""
            last_name = emp.get("last_name", "") or ""
            full_name = f"{first_name} {last_name}".strip()
            active_employees[code] = clean_txt(full_name if full_name else f"User {code}")

    # 2. Fetch daily chronological clock logs
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

    present_staff = []      
    late_staff = []         
    full_absent_staff = []  
    checkout_staff = []     
    excel_rows = [] # Compiles unified rows matching your output requirements

    for code, name in active_employees.items():
        if code in emp_punches and len(emp_punches[code]) > 0:
            user_punches = emp_punches[code]
            punch_count = len(user_punches)
            
            first_punch = user_punches.copy().pop(0)
            last_punch = user_punches[-1]
            
            clock_in_str = first_punch.strftime('%H:%M')
            
            # Determine daily tracking metrics status tags
            is_late = first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15)
            status_label = "Late(LT)" if is_late else "Present(P)"

            if first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15):
                late_staff.append((code, name, first_punch.strftime('%I:%M %p')))

            if punch_count % 2 != 0:
                present_staff.append((code, name, first_punch.strftime('%I:%M %p')))
                excel_rows.append({
                    "Employee ID": code,
                    "First Name": name,
                    "Date": selected_date_str,
                    "Clock In": clock_in_str,
                    "Clock Out": "",
                    "Total WT": "",
                    "Status": status_label
                })
            else:
                clock_out_str = last_punch.strftime('%H:%M')
                checkout_staff.append((code, name, last_punch.strftime('%I:%M %p')))
                
                # Math calculation loop processing Total Work Time (Total WT)
                time_diff = last_punch - first_punch
                total_seconds = int(time_diff.total_seconds())
                hours = total_seconds // 3600
                minutes = (total_seconds % 3600) // 60
                total_wt_str = f"{hours:02d}:{minutes:02d}"
                
                excel_rows.append({
                    "Employee ID": code,
                    "First Name": name,
                    "Date": selected_date_str,
                    "Clock In": clock_in_str,
                    "Clock Out": clock_out_str,
                    "Total WT": total_wt_str,
                    "Status": status_label
                })
        else:
            full_absent_staff.append((code, name))
            excel_rows.append({
                "Employee ID": code,
                "First Name": name,
                "Date": selected_date_str,
                "Clock In": "",
                "Clock Out": "",
                "Total WT": "",
                "Status": "Absence(A)"
            })

    full_absent_staff.sort(key=lambda val: int(val) if str(val).isdigit() else str(val))
    present_staff.sort(key=lambda val: int(val) if str(val).isdigit() else str(val))
    late_staff.sort(key=lambda val: int(val) if str(val).isdigit() else str(val))
    checkout_staff.sort(key=lambda val: int(val) if str(val).isdigit() else str(val))
    
    # Keep final exported rows cleanly sorted by numerical employee ID
    excel_rows.sort(key=lambda row_item: int(row_item["Employee ID"]) if str(row_item["Employee ID"]).isdigit() else row_item["Employee ID"])
    
    return active_employees, present_staff, late_staff, full_absent_staff, checkout_staff, excel_rows
# ==========================================
# 3. INTERFACE RENDERING & CONTROLS
# ==========================================
now_syria = datetime.now(SYRIA_TZ)
time_str = now_syria.strftime('%I:%M:%S %p')
selected_date_str = now_syria.strftime('%Y-%m-%d')

st.title(TEXT_CONFIG["title_main"])
st.markdown(TEXT_CONFIG["lbl_date"].format(selected_date_str, time_str))

btn_col1, btn_col2 = st.columns(2)
with btn_col1:
    if st.button(TEXT_CONFIG["btn_refresh"], use_container_width=True):
        st.cache_data.clear()
        st.rerun()

try:
    active_employees, present_staff, late_staff, full_absent_staff, checkout_staff, excel_rows = load_attendance_data_from_api(selected_date_str)
    
    # 📊 Unified structured DataFrame build pipeline matching your exact image layout requirements
    df_report = pd.DataFrame(excel_rows)
    csv_data = df_report.to_csv(index=False).encode('utf-8-sig')
    
    with btn_col2:
        st.download_button(
            label=TEXT_CONFIG["btn_download_excel"],
            data=csv_data,
            file_name=f"Daily_Attendance_Report_{selected_date_str}.csv",
            mime="text/csv",
            use_container_width=True
        )

    # Totals computation metrics calculations
    total_emp = len(active_employees)
    p_count = len(present_staff)
    l_count = len(late_staff)
    c_count = len(checkout_staff)
    a_count = len(full_absent_staff)
    
    st.write("### 📊 Interactive Metrics Panels (Click an item to see details right below it)")
    current_view = st.session_state["selected_view"]
    
    # 1. Total Registered Card Button + Direct Nested Viewer Dropdown Layout
    if st.button(f"👥 Total Active Staff Count ── {total_emp}"):
        st.session_state["selected_view"] = "all"
        st.rerun()
    if current_view == "all":
        st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
        st.write(f"### {TEXT_CONFIG['header_all'].format(total_emp)}")
        for code, name in sorted(active_employees.items(), key=lambda x: int(x) if x.isdigit() else x):
            st.markdown(TEXT_CONFIG["all_row"].format(name, code))
        st.markdown('</div>', unsafe_allow_html=True)
        
    # 2. On Premises Card Button + Direct Nested Viewer Dropdown Layout
    if st.button(f"🟢 Present Staff On-Premises ── {p_count}"):
        st.session_state["selected_view"] = "present"
        st.rerun()
    if current_view == "present":
        st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
        st.write(f"### {TEXT_CONFIG['header_present'].format(p_count)}")
        if present_staff:
            for code, name, time_in in present_staff:
                st.markdown(TEXT_CONFIG["present_row"].format(name, code, time_in))
        else:
            st.info("No employee logs are active on premises right now.")
        st.markdown('</div>', unsafe_allow_html=True)
        
    # 3. Late Arrival Card Button + Direct Nested Viewer Dropdown Layout
    if st.button(f"⏰ Late Arrivals Logged ── {l_count}"):
        st.session_state["selected_view"] = "late"
        st.rerun()
    if current_view == "late":
        st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
        st.write(f"### {TEXT_CONFIG['header_late'].format(l_count)}")
        if late_staff:
            for code, name, time_in in late_staff:
                st.markdown(TEXT_CONFIG["late_row"].format(name, code, time_in))
        else:
            st.success("🎉 Exceptional performance! Zero late arrivals tracked today.")
        st.markdown('</div>', unsafe_allow_html=True)
        
    # 4. Departed Workday Card Button + Direct Nested Viewer Dropdown Layout
    if st.button(f"✅ Checked-Out / Shifts Completed ── {c_count}"):
        st.session_state["selected_view"] = "checkout"
        st.rerun()
    if current_view == "checkout":
        st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
        st.write(f"### {TEXT_CONFIG['header_checkout'].format(c_count)}")
        if checkout_staff:
            for code, name, time_out in checkout_staff:
                st.markdown(TEXT_CONFIG["checkout_row"].format(name, code, time_out))
        else:
            st.info("No employee departure punch records found yet.")
        st.markdown('</div>', unsafe_allow_html=True)
        
    # 5. Absentees Card Button + Direct Nested Viewer Dropdown Layout
    if st.button(f"❌ Full Absences Recorded ── {a_count}"):
        st.session_state["selected_view"] = "absent"
        st.rerun()
    if current_view == "absent":
        st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
        st.write(f"### {TEXT_CONFIG['header_absent'].format(a_count)}")
        if full_absent_staff:
            for code, name in full_absent_staff:
                st.markdown(TEXT_CONFIG["absent_row"].format(name, code))
        else:
            st.success("🎉 Flawless operations! Perfect 100% attendance tracked today.")
        st.markdown('</div>', unsafe_allow_html=True)

except Exception as e:
    st.error(TEXT_CONFIG["err_api"].format(str(e)))
    if st.checkbox("Show System Operational Debug Logs"):
        for log in st.session_state.get("debug_logs", []):
            st.text(log)
