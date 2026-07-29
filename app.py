import streamlit as st
import requests
import unicodedata
import pandas as pd
import io
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
# 1. ROTATING SATELLITE & 3+2 MOBILE GRID CSS
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

    /* 📡 ROTATING SATELLITE DISH ANIMATION */
    .satellite-container {
        display: flex;
        align-items: center;
        justify-content: center;
        background: linear-gradient(145deg, #ffffff, #f0f4f8);
        border: 1px solid #e2e8f0;
        padding: 8px 20px;
        border-radius: 30px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.03);
        margin: 0 auto 15px auto;
        width: fit-content;
        gap: 12px;
    }
    .dish-icon {
        position: relative;
        width: 30px;
        height: 30px;
        color: #0ea5e9;
        display: flex;
        align-items: center;
        justify-content: center;
    }
    .dish-svg {
        z-index: 2;
        width: 100%;
        height: 100%;
        animation: rotate-dish 6s linear infinite;
        transform-origin: center;
    }
    @keyframes rotate-dish {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    
    .signal-waves {
        position: absolute;
        top: -4px; right: -4px;
        width: 12px; height: 12px;
    }
    .wave {
        position: absolute;
        border: 2px solid #0ea5e9;
        border-radius: 50%;
        opacity: 0;
        animation: emit-signal 2s linear infinite;
    }
    .wave:nth-child(1) { width: 8px; height: 8px; top: 8px; left: -4px; animation-delay: 0s; }
    .wave:nth-child(2) { width: 16px; height: 16px; top: 4px; left: -8px; animation-delay: 0.7s; }
    
    @keyframes emit-signal {
        0% { transform: scale(0.5); opacity: 0; }
        50% { opacity: 1; }
        100% { transform: scale(1.6); opacity: 0; }
    }
    
    .satellite-text {
        font-size: 13px;
        font-weight: 800;
        color: #0f172a;
        background: -webkit-linear-gradient(0deg, #0ea5e9, #10b981);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }

    /* KPI BUTTONS */
    div[data-testid="stColumn"] button {
        width: 100% !important;
        background: #ffffff !important;
        border-radius: 10px !important;
        padding: 10px 4px !important;
        text-align: center !important;
        box-shadow: 0 2px 5px rgba(0,0,0,0.03) !important;
        border: 1px solid #e2e8f0 !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stColumn"] button p {
        font-size: 12px !important;
        color: #1e293b !important;
        font-weight: bold !important;
        margin: 0 !important;
        white-space: pre-line !important;
        line-height: 1.3 !important;
    }

    /* 📱 MOBILE 3+2 GRID FIX FOR KPI BUTTONS */
    @media (max-width: 768px) {
        div[data-testid="stHorizontalBlock"]:has(button[key*="btn_"]) {
            display: grid !important;
            grid-template-columns: repeat(3, 1fr) !important;
            gap: 6px !important;
        }
        div[data-testid="stHorizontalBlock"]:has(button[key*="btn_"]) > div[data-testid="column"] {
            width: 100% !important;
            min-width: 0 !important;
        }
        /* Make the 4th and 5th buttons span nicely in the second row */
        div[data-testid="stHorizontalBlock"]:has(button[key*="btn_"]) > div[data-testid="column"]:nth-child(4) {
            grid-column: 1 / span 1 !important;
        }
        div[data-testid="stHorizontalBlock"]:has(button[key*="btn_"]) > div[data-testid="column"]:nth-child(5) {
            grid-column: 2 / span 2 !important;
        }
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
        padding: 10px; border-bottom: 1px solid #f1f5f9; text-align: center; font-weight: 500;
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

def load_attendance_data_from_api(selected_date_str, selected_date_obj):
    token = get_auth_token()
    if not token: raise Exception("رمز المصادقة غير صالح.")
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    
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
            # FIX: Safely parse names to prevent literal "None" string
            f_name = emp.get("first_name")
            l_name = emp.get("last_name")
            f_str = str(f_name).strip() if f_name and str(f_name).lower() != "none" else ""
            l_str = str(l_name).strip() if l_name and str(l_name).lower() != "none" else ""
            full_name = f"{f_str} {l_str}".strip()
            
            active_employees[cleaned_code] = clean_txt(full_name if full_name else f"موظف {cleaned_code}")

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
                emp_punches.setdefault(cleaned_code, []).append(p_time)
            except Exception: continue

    present_staff, late_staff, absent_staff, checkout_staff, excel_rows = [], [], [], [], []

    for code, name in active_employees.items():
        punches = sorted(emp_punches.get(code, []))
        filtered_punches = []
        for p in punches:
            if not filtered_punches or abs((p - filtered_punches[-1]).total_seconds()) > 60:
                filtered_punches.append(p)

        day_punches = [p for p in filtered_punches if p.date() == selected_date_obj and p.hour >= 5]
        
        if not day_punches:
            absent_staff.append((code, name))
            excel_rows.append({"ID": code, "Name": name, "Status": "Absent"})
            continue

        first_p = day_punches[0]
        is_late = first_p.hour > 9 or (first_p.hour == 9 and first_p.minute > 15)
        if is_late: late_staff.append((code, name, first_p.strftime('%I:%M %p')))

        next_morning = [p for p in filtered_punches if p.date() == selected_date_obj + timedelta(days=1) and p.hour < 5]
        punch_count = 2 if (len(day_punches) % 2 != 0 and next_morning) else len(day_punches)

        if punch_count % 2 != 0:
            present_staff.append((code, name, first_p.strftime('%I:%M %p')))
            excel_rows.append({"ID": code, "Name": name, "Status": "Present"})
        else:
            last_p = (next_morning[-1] if (len(day_punches) % 2 != 0 and next_morning) else day_punches[-1])
            checkout_staff.append((code, name, last_p.strftime('%I:%M %p')))
            excel_rows.append({"ID": code, "Name": name, "Status": "Checkout"})

    absent_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else 999)
    present_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else 999)
    return active_employees, present_staff, late_staff, absent_staff, checkout_staff, excel_rows

# ==========================================
# 3. INTERFACE RENDERING
# ==========================================
now_syria = datetime.now(SYRIA_TZ)

# 📡 ROTATING SATELLITE HEADER
st.markdown(
    """
    <div class="satellite-container">
        <div class="dish-icon">
            <svg class="dish-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M4 10a7.31 7.31 0 0 0 10 10Z"/>
                <path d="m9 15-1.5 1.5"/>
                <path d="M12 12a1 1 0 0 0-1-1 1 1 0 0 0-1 1 1 1 0 0 0 1 1 1 1 0 0 0 1-1Z"/>
                <path d="M20 4a10.82 10.82 0 0 0-10 10"/>
                <path d="m16 8-1.5 1.5"/>
                <path d="M18.8 6.2C19.5 7.1 20 8.3 20 9.5"/>
                <path d="M21 4c.6.9 1 2 1 3.2"/>
            </svg>
            <div class="signal-waves">
                <div class="wave"></div>
                <div class="wave"></div>
            </div>
        </div>
        <span class="satellite-text">الرادار متصل ويقوم بالمسح الحي</span>
    </div>
    """, unsafe_allow_html=True
)

c_date, c_ref = st.columns(2)
with c_date:
    selected_date_str = st.date_input("", value=now_syria.date(), label_visibility="collapsed").strftime('%Y-%m-%d')
with c_ref:
    if st.button("🔄 تحديث", use_container_width=True): st.cache_data.clear(); st.rerun()

try:
    act, pre, lat, abs_s, chk, exc = load_attendance_data_from_api(selected_date_str, datetime.strptime(selected_date_str, "%Y-%m-%d").date())
    
    # 📱 5 KPI BUTTONS IN A SINGLE CONTAINER (Automatically wraps 3 + 2 on mobile via CSS)
    b1, b2, b3, b4, b5 = st.columns(5)
    with b1:
        if st.button(f"👥 الكل\n\n{len(act)}", key="btn_all"): st.session_state["selected_view"] = "all"; st.rerun()
    with b2:
        if st.button(f"🟢 بالعمل\n\n{len(pre)}", key="btn_pre"): st.session_state["selected_view"] = "present"; st.rerun()
    with b3:
        if st.button(f"⏰ متأخر\n\n{len(lat)}", key="btn_lat"): st.session_state["selected_view"] = "late"; st.rerun()
    with b4:
        if st.button(f"🏁 انصراف\n\n{len(chk)}", key="btn_chk"): st.session_state["selected_view"] = "checkout"; st.rerun()
    with b5:
        if st.button(f"❌ غياب\n\n{len(abs_s)}", key="btn_abs"): st.session_state["selected_view"] = "absent"; st.rerun()

    search_query = st.text_input("", placeholder=TEXT_CONFIG["search_placeholder"], label_visibility="collapsed").strip().lower()
    match = lambda c, n: (search_query in str(c).lower() or search_query in str(n).lower()) if search_query else True

    view = st.session_state["selected_view"]

    if view == "all":
        rows = [f"<tr><td>{c}</td><td>{n}</td><td><span class='badge-present'>نشط</span></td></tr>" for c, n in act.items() if match(c, n)]
        st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="3" class="table-main-title-header">{TEXT_CONFIG["header_all"].format(len(act))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الحالة</th></tr>{"".join(rows)}</table>', unsafe_allow_html=True)
        
    elif view == "present":
        if pre:
            rows = [f"<tr><td>{c}</td><td>{n}</td><td>{t}</td></tr>" for c, n, t in pre if match(c, n)]
            st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="3" class="table-main-title-header">{TEXT_CONFIG["header_present"].format(len(pre))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الدخول</th></tr>{"".join(rows)}</table>', unsafe_allow_html=True)
        else: st.info("لا يوجد موظفين متواجدين حالياً.")
        
    elif view == "late":
        if lat:
            rows = [f"<tr><td>{c}</td><td>{n}</td><td>{t}</td><td><span class='badge-late'>متأخر</span></td></tr>" for c, n, t in lat if match(c, n)]
            st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="4" class="table-main-title-header">{TEXT_CONFIG["header_late"].format(len(lat))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الدخول</th><th>الحالة</th></tr>{"".join(rows)}</table>', unsafe_allow_html=True)
        else: st.success("لا يوجد متأخرين!")
        
    elif view == "checkout":
        if chk:
            rows = [f"<tr><td>{c}</td><td>{n}</td><td>{t}</td></tr>" for c, n, t in chk if match(c, n)]
            st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="3" class="table-main-title-header">{TEXT_CONFIG["header_checkout"].format(len(chk))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الانصراف</th></tr>{"".join(rows)}</table>', unsafe_allow_html=True)
        else: st.info("لا توجد عمليات انصراف مسجلة.")
        
    elif view == "absent":
        if abs_s:
            rows = [f"<tr><td>{c}</td><td>{n}</td><td><span class='badge-absent'>غياب</span></td></tr>" for c, n in abs_s if match(c, n)]
            st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="3" class="table-main-title-header">{TEXT_CONFIG["header_absent"].format(len(abs_s))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الحالة</th></tr>{"".join(rows)}</table>', unsafe_allow_html=True)

except Exception as e:
    st.error(TEXT_CONFIG["err_api"].format(str(e)))
