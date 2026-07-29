import streamlit as st
import requests
import unicodedata
import pandas as pd
import io
import base64
from datetime import datetime, timedelta
import zoneinfo

from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 0. RTL ARABIC TEXT & VISUAL CONFIG
# ==========================================
TEXT_CONFIG = {
    "page_title": "حضور وانصراف القصر الذهبي",
    "title_main": "✨ شركة القصر الذهبي ✨",
    "search_placeholder": "🔍 ابحث باسم الموظف أو رقم الكود...",
    
    "header_all": "👥 كافة موظفي الشركة النشطين ({})",
    "header_present": "🟢 المتواجدون حالياً ({})",
    "header_late": "⏰ المتأخرون ({})",
    "header_checkout": "🏁 المنصرفون ({})",
    "header_absent": "❌ الغيابات ({})",
    "err_api": "خطأ في الاتصال بواجهة BioTime: {}"
}

st.set_page_config(page_title=TEXT_CONFIG["page_title"], page_icon="📡", layout="wide", initial_sidebar_state="collapsed")

# ==========================================
# 1. 360 FULL ROTATION DISH ANIMATION & CSS
# ==========================================
st.markdown("""
    <style>
    /* Clean UI */
    header[data-testid="stHeader"] { display: none !important; }
    footer { display: none !important; }
    .stApp { direction: rtl; background-color: #f4f7f9; font-family: system-ui, -apple-system, sans-serif; }
    
    .block-container {
        padding-top: 15px !important;
        padding-bottom: 30px !important; 
        padding-left: 10px !important;
        padding-right: 10px !important;
        max-width: 100% !important;
    }

    /* 📡 360 DEGREE CONTINUOUS ROTATION HEADER */
    .status-badge {
        display: flex;
        align-items: center;
        justify-content: center;
        background: linear-gradient(145deg, #ffffff, #f0f4f8);
        border: 1px solid #cbd5e1;
        padding: 8px 20px;
        border-radius: 30px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.04);
        margin: 0 auto 15px auto;
        width: fit-content;
        gap: 12px;
    }
    
    .animated-dish {
        width: 34px;
        height: 34px;
        object-fit: contain;
        transform-origin: center center;
        animation: rotate-360 5s linear infinite;
    }
    
    @keyframes rotate-360 {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }

    .status-indicator {
        display: flex;
        align-items: center;
        gap: 8px;
    }

    .blinking-dot {
        width: 10px;
        height: 10px;
        background-color: #22c55e;
        border-radius: 50%;
        display: inline-block;
        box-shadow: 0 0 0 0 rgba(34, 197, 94, 0.7);
        animation: pulse-green 1.5s infinite;
    }

    @keyframes pulse-green {
        0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(34, 197, 94, 0.7); }
        70% { transform: scale(1); box-shadow: 0 0 0 6px rgba(34, 197, 94, 0); }
        100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(34, 197, 94, 0); }
    }

    .online-text {
        font-size: 14px;
        font-weight: 800;
        color: #0f172a;
        letter-spacing: 0.5px;
    }

    /* 📱 WIDE & CENTERED MOBILE BUTTONS */
    div[data-testid="stColumn"] button {
        width: 100% !important;
        background: #ffffff !important;
        border-radius: 12px !important;
        padding: 12px 10px !important;
        text-align: center !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.04) !important;
        border: 1px solid #cbd5e1 !important;
        margin-bottom: 6px !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stColumn"] button p {
        font-size: 13px !important;
        color: #1e293b !important;
        font-weight: 700 !important;
        margin: 0 !important;
        white-space: pre-line !important;
        line-height: 1.4 !important;
    }

    /* TABLE DESIGN */
    .responsive-grid-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 15px;
        font-size: 13px;
        background-color: #ffffff;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 4px 6px rgba(0,0,0,0.02);
    }
    .responsive-grid-table .table-main-title-header {
        background: #0f172a;
        color: #ffffff !important;
        text-align: center;
        font-size: 14px;
        padding: 12px;
    }
    .responsive-grid-table th {
        background-color: #f8fafc;
        color: #475569;
        padding: 10px;
        border-bottom: 1px solid #e2e8f0;
        text-align: center;
    }
    .responsive-grid-table td { 
        padding: 10px; border-bottom: 1px solid #f1f5f9; text-align: center; font-weight: 500; color: #1e293b;
    }
    .badge-present { background-color: #dcfce7; color: #166534; padding: 4px 8px; border-radius: 6px; font-size: 11px; }
    .badge-late { background-color: #fef3c7; color: #9a3412; padding: 4px 8px; border-radius: 6px; font-size: 11px; }
    .badge-absent { background-color: #fee2e2; color: #991b1b; padding: 4px 8px; border-radius: 6px; font-size: 11px; }
    </style>
""", unsafe_allow_html=True)

EXCLUDED_MANAGEMENT_CODES = ("40",)
EXCLUDED_RESIGNED_CODES = ("34",) 
SYRIA_TZ = zoneinfo.ZoneInfo("Asia/Damascus")

BASE_URL = st.secrets["biotime"]["base_url"].rstrip('/')
TOKEN_URL = st.secrets["biotime"]["token_url"]
EMAIL = st.secrets["biotime"]["email"]
PASSWORD = st.secrets["biotime"]["password"]
COMPANY = st.secrets["biotime"]["company"]

if "debug_logs" not in st.session_state: st.session_state["debug_logs"] = []
if "selected_view" not in st.session_state: st.session_state["selected_view"] = "present"

def clean_txt(raw_text): return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip()) if raw_text else ""

@st.cache_data(ttl=300)
def get_auth_token():
    try:
        res = requests.post(TOKEN_URL, json={"email": EMAIL, "password": PASSWORD, "company": COMPANY}, timeout=10)
        if res.status_code in (200, 201): return res.json().get("token")
    except Exception: return None

@st.cache_data(ttl=120)
def get_biometric_devices():
    token = get_auth_token()
    if not token: return []
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    for endpoint in ["/iclock/api/terminals/", "/iclock/api/devices/"]:
        try:
            res = requests.get(f"{BASE_URL}{endpoint}", headers=headers, timeout=10)
            if res.status_code == 200:
                data = res.json()
                return data.get("data", data) if isinstance(data, (dict, list)) else []
        except Exception:
            continue
    return []

def load_attendance_data_from_api(selected_date_str, selected_date_obj):
    token = get_auth_token()
    if not token: raise Exception("رمز المصادقة غير صالح.")
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    
    # 1. Fetch Devices & Build Mapping
    devices = []
    try:
        for endpoint in ["/iclock/api/terminals/", "/iclock/api/devices/"]:
            dev_res = requests.get(f"{BASE_URL}{endpoint}", headers=headers, timeout=10)
            if dev_res.status_code == 200:
                d_data = dev_res.json()
                devices = d_data.get("data", d_data) if isinstance(d_data, (dict, list)) else []
                break
    except Exception:
        pass
    
    terminal_map = {}
    for d in devices:
        sn = str(d.get("sn", ""))
        alias = d.get("alias") or d.get("terminal_name") or sn
        if sn:
            terminal_map[sn] = alias

    # 2. Fetch Employees
    all_employees = []
    try:
        emp_res = requests.get(f"{BASE_URL}/personnel/api/employees/?page_size=1000", headers=headers, timeout=15)
        if emp_res.status_code == 200: all_employees = emp_res.json().get("data", [])
    except Exception: pass

    active_employees = {}
    for emp in all_employees:
        raw_code = str(emp.get("emp_code", "")).strip()
        cleaned_code = str(int(raw_code)) if raw_code.isdigit() else raw_code
        if cleaned_code and cleaned_code not in EXCLUDED_MANAGEMENT_CODES and cleaned_code not in EXCLUDED_RESIGNED_CODES:
            f_name = emp.get("first_name")
            l_name = emp.get("last_name")
            f_str = str(f_name).strip() if f_name and str(f_name).lower() != "none" else ""
            l_str = str(l_name).strip() if l_name and str(l_name).lower() != "none" else ""
            full_name = f"{f_str} {l_str}".strip()
            
            active_employees[cleaned_code] = clean_txt(full_name if full_name else f"موظف {cleaned_code}")

    # 3. Fetch Transactions Logs
    prev_day = (selected_date_obj - timedelta(days=1)).strftime('%Y-%m-%d') + " 00:00:00"
    next_day = (selected_date_obj + timedelta(days=1)).strftime('%Y-%m-%d') + " 05:00:00"
    
    raw_logs = []
    try:
        logs_res = requests.get(f"{BASE_URL}/iclock/api/transactions/?start_time={prev_day}&end_time={next_day}&page_size=5000", headers=headers, timeout=15)
        if logs_res.status_code == 200: raw_logs = logs_res.json().get("data", [])
    except Exception: pass

    emp_punches = {}
    for log in raw_logs:
        raw_code = str(log.get("emp_code", "")).strip()
        cleaned_code = str(int(raw_code)) if raw_code.isdigit() else raw_code
        if cleaned_code in active_employees and log.get("punch_time"):
            try:
                p_time = datetime.strptime(log.get("punch_time")[:19], "%Y-%m-%d %H:%M:%S")
                dev_sn = str(log.get("terminal_sn", ""))
                dev_name = log.get("terminal_alias") or log.get("terminal_name") or terminal_map.get(dev_sn, dev_sn or "جهاز رئيسي")
                emp_punches.setdefault(cleaned_code, []).append((p_time, dev_name))
            except Exception: continue

    present_staff, late_staff, absent_staff, checkout_staff, excel_rows = [], [], [], [], []

    for code, name in active_employees.items():
        punches = sorted(emp_punches.get(code, []), key=lambda x: x[0])
        filtered_punches = []
        for p_time, d_name in punches:
            if not filtered_punches or abs((p_time - filtered_punches[-1][0]).total_seconds()) > 60:
                filtered_punches.append((p_time, d_name))

        day_punches = [(p, d) for p, d in filtered_punches if p.date() == selected_date_obj and p.hour >= 5]
        
        if not day_punches:
            absent_staff.append((code, name))
            excel_rows.append({"ID": code, "Name": name, "Status": "Absent"})
            continue

        first_p, first_dev = day_punches[0]
        is_late = first_p.hour > 9 or (first_p.hour == 9 and first_p.minute > 15)
        if is_late: late_staff.append((code, name, first_p.strftime('%I:%M %p'), first_dev))

        next_morning = [(p, d) for p, d in filtered_punches if p.date() == selected_date_obj + timedelta(days=1) and p.hour < 5]
        punch_count = 2 if (len(day_punches) % 2 != 0 and next_morning) else len(day_punches)

        if punch_count % 2 != 0:
            present_staff.append((code, name, first_p.strftime('%I:%M %p'), first_dev))
            excel_rows.append({"ID": code, "Name": name, "Status": "Present"})
        else:
            last_p, last_dev = (next_morning[-1] if (len(day_punches) % 2 != 0 and next_morning) else day_punches[-1])
            checkout_staff.append((code, name, last_p.strftime('%I:%M %p'), last_dev))
            excel_rows.append({"ID": code, "Name": name, "Status": "Checkout"})

    absent_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else 999)
    present_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else 999)
    return active_employees, present_staff, late_staff, absent_staff, checkout_staff, devices, excel_rows

# ==========================================
# 3. INTERFACE RENDERING
# ==========================================
now_syria = datetime.now(SYRIA_TZ)

dish_img_tag = ""
try:
    with open("image_632b3d.jpg", "rb") as img_file:
        encoded_string = base64.b64encode(img_file.read()).decode()
        dish_img_tag = f'<img src="data:image/jpeg;base64,{encoded_string}" class="animated-dish" />'
except Exception:
    dish_img_tag = '<div class="animated-dish" style="font-size: 24px;">📡</div>'

# 📡 360 ROTATING DISH & ONLINE STATUS HEADER
st.markdown(
    f"""
    <div class="status-badge">
        {dish_img_tag}
        <div class="status-indicator">
            <span class="blinking-dot"></span>
            <span class="online-text">Online</span>
        </div>
    </div>
    """, unsafe_allow_html=True
)

c_date, c_ref = st.columns(2)
with c_date:
    selected_date_str = st.date_input("", value=now_syria.date(), label_visibility="collapsed").strftime('%Y-%m-%d')
with c_ref:
    if st.button("🔄 تحديث البيانات", use_container_width=True): st.cache_data.clear(); st.rerun()

try:
    act, pre, lat, abs_s, chk, devices, exc = load_attendance_data_from_api(selected_date_str, datetime.strptime(selected_date_str, "%Y-%m-%d").date())
    
    # 📱 WIDE & CENTERED MOBILE BUTTONS LAYOUT
    if st.button(f"👥 كافة موظفي الشركة النشطين ({len(act)})", use_container_width=True): 
        st.session_state["selected_view"] = "all"
    
    col_p, col_l = st.columns(2)
    with col_p:
        if st.button(f"🟢 المتواجدون ({len(pre)})", use_container_width=True): 
            st.session_state["selected_view"] = "present"
    with col_l:
        if st.button(f"⏰ المتأخرون ({len(lat)})", use_container_width=True): 
            st.session_state["selected_view"] = "late"

    col_c, col_a = st.columns(2)
    with col_c:
        if st.button(f"🏁 المنصرفون ({len(chk)})", use_container_width=True): 
            st.session_state["selected_view"] = "checkout"
    with col_a:
        if st.button(f"❌ الغيابات ({len(abs_s)})", use_container_width=True): 
            st.session_state["selected_view"] = "absent"

    # 🖨️ BIOMETRIC DEVICES EXPANDER
    with st.expander("🖨️ أجهزة الحضور والانصراف المرتبطة", expanded=False):
        if devices:
            dev_rows = []
            for d in devices:
                d_name = d.get("alias") or d.get("terminal_name") or d.get("sn", "جهاز غير محدد")
                d_sn = d.get("sn", "N/A")
                d_ip = d.get("ip_address", "غير متوفر")
                dev_rows.append(f"<tr><td>{d_name}</td><td>{d_sn}</td><td>{d_ip}</td><td><span class='badge-present'>متصل</span></td></tr>")
            st.markdown(f'<table class="responsive-grid-table"><tr><th>اسم الجهاز</th><th>الرقم التسلسلي (SN)</th><th>عنوان IP</th><th>الحالة</th></tr>{"".join(dev_rows)}</table>', unsafe_allow_html=True)
        else:
            st.info("لا توجد أجهزة مسجلة حالياً أو تعذر جلبها.")

    search_query = st.text_input("", placeholder=TEXT_CONFIG["search_placeholder"], label_visibility="collapsed").strip().lower()
    match = lambda c, n: (search_query in str(c).lower() or search_query in str(n).lower()) if search_query else True

    view = st.session_state["selected_view"]

    if view == "all":
        rows = [f"<tr><td>{c}</td><td>{n}</td><td><span class='badge-present'>نشط</span></td></tr>" for c, n in act.items() if match(c, n)]
        st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="3" class="table-main-title-header">{TEXT_CONFIG["header_all"].format(len(act))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الحالة</th></tr>{"".join(rows)}</table>', unsafe_allow_html=True)
        
    elif view == "present":
        if pre:
            rows = [f"<tr><td>{c}</td><td>{n}</td><td>{t}</td><td>{d}</td></tr>" for c, n, t, d in pre if match(c, n)]
            st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="4" class="table-main-title-header">{TEXT_CONFIG["header_present"].format(len(pre))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الدخول</th><th>جهاز البصمة</th></tr>{"".join(rows)}</table>', unsafe_allow_html=True)
        else: st.info("لا يوجد موظفين متواجدين حالياً.")
        
    elif view == "late":
        if lat:
            rows = [f"<tr><td>{c}</td><td>{n}</td><td>{t}</td><td><span class='badge-late'>متأخر</span></td><td>{d}</td></tr>" for c, n, t, d in lat if match(c, n)]
            st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="5" class="table-main-title-header">{TEXT_CONFIG["header_late"].format(len(lat))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الدخول</th><th>الحالة</th><th>جهاز البصمة</th></tr>{"".join(rows)}</table>', unsafe_allow_html=True)
        else: st.success("لا يوجد متأخرين!")
        
    elif view == "checkout":
        if chk:
            rows = [f"<tr><td>{c}</td><td>{n}</td><td>{t}</td><td>{d}</td></tr>" for c, n, t, d in chk if match(c, n)]
            st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="4" class="table-main-title-header">{TEXT_CONFIG["header_checkout"].format(len(chk))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الانصراف</th><th>جهاز البصمة</th></tr>{"".join(rows)}</table>', unsafe_allow_html=True)
        else: st.info("لا توجد عمليات انصراف مسجلة.")
        
    elif view == "absent":
        if abs_s:
            rows = [f"<tr><td>{c}</td><td>{n}</td><td><span class='badge-absent'>غياب</span></td></tr>" for c, n in abs_s if match(c, n)]
            st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="3" class="table-main-title-header">{TEXT_CONFIG["header_absent"].format(len(abs_s))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الحالة</th></tr>{"".join(rows)}</table>', unsafe_allow_html=True)

except Exception as e:
    st.error(TEXT_CONFIG["err_api"].format(str(e)))
