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
    "lbl_date": "📅 التاريخ للتقرير: **{}**  │  ⏰ الوقت الحالي: **{}**",
    "btn_refresh": "🔄 تحديث",
    "lbl_pick_date": "📅 اختر تاريخ التقرير:",
    "btn_download_excel": "📥 تحميل تقرير Excel النمطي",
    "search_placeholder": "🔍 ابحث باسم الموظف أو رقم الكود...",
    
    "header_all": "👥 كافة موظفي الشركة النشطين ({})",
    "header_present": "🟢 الموظفون المتواجدون حالياً ({})",
    "header_late": "⏰ الموظفون المتأخرون ({})",
    "header_checkout": "🏁 الموظفون المنصرفون ({})",
    "header_absent": "❌ قائمة الغيابات الكاملة ({})",
    "err_api": "خطأ في الاتصال بواجهة BioTime السحابية: {}"
}

st.set_page_config(page_title=TEXT_CONFIG["page_title"], page_icon="📊", layout="wide", initial_sidebar_state="collapsed")

# ==========================================
# 1. LUXURY RADAR ANIMATION & RESPONSIVE CSS
# ==========================================
st.markdown("""
    <style>
    /* Clean Page Eraser */
    header[data-testid="stHeader"] { display: none !important; }
    footer { display: none !important; }
    div[data-testid="stDecoration"] { display: none !important; }
    .stApp { direction: rtl; background-color: #f1f5f9; font-family: system-ui, -apple-system, sans-serif; }
    
    .block-container {
        padding-top: 12px !important;
        padding-bottom: 30px !important; 
        padding-left: 10px !important;
        padding-right: 10px !important;
        max-width: 100% !important;
    }

    /* 🟢 REAL DYNAMIC RADAR SWEEP ANIMATION */
    .radar-status-bar {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 10px;
        background: linear-gradient(135deg, #0f172a, #1e293b);
        color: #38bdf8;
        padding: 6px 16px;
        border-radius: 24px;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.25);
        margin: 0 auto 14px auto;
        width: fit-content;
        border: 1px solid rgba(56, 189, 248, 0.3);
    }
    .radar-dish {
        position: relative;
        width: 22px;
        height: 22px;
        border-radius: 50%;
        background: rgba(16, 185, 129, 0.15);
        border: 1.5px solid #10b981;
        overflow: hidden;
    }
    .radar-ring {
        position: absolute;
        top: 50%; left: 50%;
        transform: translate(-50%, -50%);
        width: 10px; height: 10px;
        border: 1px dashed rgba(16, 185, 129, 0.6);
        border-radius: 50%;
    }
    .radar-sweep {
        position: absolute;
        top: 0; left: 0; width: 100%; height: 100%;
        border-radius: 50%;
        background: conic-gradient(from 0deg, rgba(16, 185, 129, 0.9) 0deg, rgba(16, 185, 129, 0.05) 90deg, transparent 90deg);
        animation: radar-spin 1.4s linear infinite;
    }
    .radar-dot {
        position: absolute;
        top: 4px; right: 5px;
        width: 4px; height: 4px;
        background-color: #34d399;
        border-radius: 50%;
        box-shadow: 0 0 6px #34d399;
        animation: blip 1.4s ease-in-out infinite;
    }
    @keyframes radar-spin {
        from { transform: rotate(0deg); }
        to { transform: rotate(360deg); }
    }
    @keyframes blip {
        0%, 100% { opacity: 0.2; transform: scale(0.8); }
        50% { opacity: 1; transform: scale(1.3); }
    }
    .radar-text {
        font-size: 12px;
        font-weight: 700;
        color: #e2e8f0;
        letter-spacing: 0.3px;
    }

    /* CARD BUTTONS STYLING */
    div[data-testid="stColumn"] button {
        width: 100% !important;
        background: #ffffff !important;
        border-radius: 12px !important;
        padding: 10px 4px !important;
        text-align: center !important;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05) !important;
        border: 1px solid #cbd5e1 !important;
        transition: all 0.2s ease-in-out !important;
    }
    div[data-testid="stColumn"] button:active {
        transform: scale(0.97);
    }
    div[data-testid="stColumn"] button p {
        font-size: 13px !important;
        color: #1e293b !important;
        font-weight: 700 !important;
        margin: 0 !important;
        white-space: pre-line !important;
        line-height: 1.3 !important;
    }

    /* 📱 RESPONSIVE MOBILE GRID FIX (PREVENTS SQUEEZING & TIGHTNESS) */
    @media (max-width: 650px) {
        /* Re-layout the 5 columns into a spacious 3 + 2 Grid */
        div[data-testid="stHorizontalBlock"]:has(button[key*="kpi_"]) {
            display: grid !important;
            grid-template-columns: repeat(3, 1fr) !important;
            gap: 8px !important;
        }
        div[data-testid="stHorizontalBlock"]:has(button[key*="kpi_"]) > div[data-testid="column"]:nth-child(4) {
            grid-column: 1 / span 1 !important;
        }
        div[data-testid="stHorizontalBlock"]:has(button[key*="kpi_"]) > div[data-testid="column"]:nth-child(5) {
            grid-column: 2 / span 2 !important;
        }

        /* Prevent button squishing */
        div[data-testid="stColumn"] button {
            min-height: 56px !important;
            padding: 8px 4px !important;
        }
        div[data-testid="stColumn"] button p {
            font-size: 12px !important;
        }
    }

    /* TABLE DESIGN */
    .badge {
        padding: 4px 8px;
        border-radius: 6px;
        font-size: 11px;
        font-weight: bold;
    }
    .badge-present { background-color: #dcfce7; color: #166534; }
    .badge-late { background-color: #fef3c7; color: #9a3412; }
    .badge-absent { background-color: #fee2e2; color: #991b1b; }
    
    .responsive-grid-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 12px;
        font-size: 13px;
        background-color: #ffffff;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    .responsive-grid-table .table-main-title-header {
        background: #1e3a8a;
        color: #ffffff !important;
        text-align: center;
        font-size: 14px;
        font-weight: bold;
        padding: 12px;
    }
    .responsive-grid-table th {
        background-color: #f8fafc;
        color: #475569;
        padding: 10px 6px;
        font-weight: 700;
        border-bottom: 1px solid #e2e8f0;
        text-align: center;
    }
    .responsive-grid-table td { 
        padding: 10px 6px; 
        border-bottom: 1px solid #f1f5f9; 
        color: #334155; 
        text-align: center;
    }
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

def log_debug(message): st.session_state["debug_logs"].append(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
def clean_txt(raw_text):
    if not raw_text: return ""
    return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip())

@st.cache_data(ttl=300)
def get_auth_token():
    payload = {"email": EMAIL, "password": PASSWORD, "company": COMPANY}
    try:
        response = requests.post(TOKEN_URL, json=payload, timeout=10)
        if response.status_code in (200, 201): return response.json().get("token")
        return None
    except Exception: return None

def load_attendance_data_from_api(selected_date_str, selected_date_obj):
    token = get_auth_token()
    if not token: raise Exception("رمز المصادقة غير صالح.")
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    
    emp_url = f"{BASE_URL}/personnel/api/employees/?page_size=1000"
    all_employees = []
    try:
        emp_res = requests.get(emp_url, headers=headers, timeout=15)
        if emp_res.status_code == 200: all_employees = emp_res.json().get("data", [])
    except Exception as e: log_debug(f"Employee Request Error: {str(e)}")

    active_employees = {}
    for emp in all_employees:
        raw_code = str(emp.get("emp_code", "")).strip()
        if not raw_code: continue
        cleaned_code = str(int(raw_code)) if raw_code.isdigit() else raw_code
        if cleaned_code not in EXCLUDED_MANAGEMENT_CODES and cleaned_code not in EXCLUDED_RESIGNED_CODES:
            first_name = emp.get("first_name", "") or ""
            last_name = emp.get("last_name", "") or ""
            full_name = f"{first_name} {last_name}".strip()
            active_employees[cleaned_code] = clean_txt(full_name if full_name else f"User {cleaned_code}")

    prev_day_obj = selected_date_obj - timedelta(days=1)
    next_day_obj = selected_date_obj + timedelta(days=1)
    
    start_window_str = prev_day_obj.strftime('%Y-%m-%d') + " 00:00:00"
    end_window_str = next_day_obj.strftime('%Y-%m-%d') + " 05:00:00"
    
    logs_url = f"{BASE_URL}/iclock/api/transactions/?start_time={start_window_str}&end_time={end_window_str}&page_size=5000"
    raw_logs = []
    try:
        logs_res = requests.get(logs_url, headers=headers, timeout=15)
        if logs_res.status_code == 200: raw_logs = logs_res.json().get("data", [])
    except Exception as e: log_debug(f"Logs Request Error: {str(e)}")

    emp_punches = {}
    for log in raw_logs:
        raw_code = str(log.get("emp_code", "")).strip()
        if not raw_code: continue
        cleaned_code = str(int(raw_code)) if raw_code.isdigit() else raw_code
        if cleaned_code in active_employees:
            punch_time_str = log.get("punch_time", "")
            terminal_alias = clean_txt(log.get("terminal_alias", "") or log.get("terminal_sn", "") or "جهاز البصمة الرئيسي")
            
            if punch_time_str:
                try:
                    p_time = datetime.strptime(punch_time_str[:19], "%Y-%m-%d %H:%M:%S")
                    if cleaned_code not in emp_punches: emp_punches[cleaned_code] = []
                    emp_punches[cleaned_code].append((p_time, terminal_alias))
                except Exception: continue

    present_staff, late_staff, full_absent_staff, checkout_staff, excel_rows = [], [], [], [], []

    for code, name in active_employees.items():
        raw_user_punches = sorted(emp_punches.get(code, []), key=lambda item: item[0])
        
        user_all_punches = []
        for current in raw_user_punches:
            if not user_all_punches: user_all_punches.append(current)
            else:
                last_saved = user_all_punches[-1]
                if abs((current[0] - last_saved[0]).total_seconds()) < 61: continue
                user_all_punches.append(current)

        cleaned_current_day_punches = []
        day_raw_punches = [item for item in user_all_punches if item[0].date() == selected_date_obj]
        
        for item in day_raw_punches:
            p_time, dev_name = item
            if p_time.hour < 5: continue
            cleaned_current_day_punches.append(item)

        if not cleaned_current_day_punches:
            full_absent_staff.append((code, name))
            excel_rows.append({
                "Employee ID": code, "First Name": name, "Date": selected_date_str,
                "Clock In": "", "Clock Out": "", "Total WT": "", "Status": "Absence(A)", "Device": "-"
            })
            continue

        first_punch_obj, first_device = cleaned_current_day_punches[0]
        clock_in_str = first_punch_obj.strftime('%H:%M')
        
        is_late = first_punch_obj.hour > 9 or (first_punch_obj.hour == 9 and first_punch_obj.minute > 15)
        status_label = "Late(LT)" if is_late else "Present(P)"
        
        if is_late: late_staff.append((code, name, first_punch_obj.strftime('%I:%M %p'), first_device))

        early_morning_punches_next_day = [item for item in user_all_punches if item[0].date() == next_day_obj and item[0].hour < 5]
        if len(cleaned_current_day_punches) % 2 != 0 and early_morning_punches_next_day:
            last_punch_obj, last_device = early_morning_punches_next_day[-1]
            punch_count = 2
        else:
            last_punch_obj, last_device = cleaned_current_day_punches[-1]
            punch_count = len(cleaned_current_day_punches)

        if punch_count % 2 != 0:
            present_staff.append((code, name, first_punch_obj.strftime('%I:%M %p'), first_device))
            excel_rows.append({
                "Employee ID": code, "First Name": name, "Date": selected_date_str,
                "Clock In": clock_in_str, "Clock Out": "", "Total WT": "", "Status": status_label, "Device": first_device
            })
        else:
            clock_out_str = last_punch_obj.strftime('%H:%M')
            checkout_staff.append((code, name, last_punch_obj.strftime('%I:%M %p'), last_device))
            
            time_diff = last_punch_obj - first_punch_obj
            total_seconds = int(time_diff.total_seconds())
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            total_wt_str = f"{hours:02d}:{minutes:02d}"
            
            excel_rows.append({
                "Employee ID": code, "First Name": name, "Date": selected_date_str,
                "Clock In": clock_in_str, "Clock Out": clock_out_str, "Total WT": total_wt_str, "Status": status_label, "Device": last_device
            })

    full_absent_staff.sort(key=lambda val: int(val[0]) if val[0].isdigit() else 999)
    present_staff.sort(key=lambda val: int(val[0]) if val[0].isdigit() else 999)
    late_staff.sort(key=lambda val: int(val[0]) if val[0].isdigit() else 999)
    checkout_staff.sort(key=lambda val: int(val[0]) if val[0].isdigit() else 999)
    excel_rows.sort(key=lambda row_item: int(row_item["Employee ID"]) if str(row_item["Employee ID"]).isdigit() else 999)
    
    return active_employees, present_staff, late_staff, full_absent_staff, checkout_staff, excel_rows

# ==========================================
# 3. INTERFACE RENDERING & CONTROLS MATRIX
# ==========================================
now_syria = datetime.now(SYRIA_TZ)

# 🔥 LIVE ANIMATED RADAR SCOPE BAR
st.markdown(
    """
    <div class="radar-status-bar">
        <div class="radar-dish">
            <div class="radar-ring"></div>
            <div class="radar-sweep"></div>
            <div class="radar-dot"></div>
        </div>
        <span class="radar-text">نظام المراقبة الحية نشط │ Live Radar</span>
    </div>
    """, unsafe_allow_html=True
)

col_left, col_right = st.columns(2)
with col_left:
    selected_date = st.date_input("", value=now_syria.date(), label_visibility="collapsed")
    selected_date_str = selected_date.strftime('%Y-%m-%d')
with col_right:
    if st.button("🔄 تحديث", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

try:
    active_employees, present_staff, late_staff, full_absent_staff, checkout_staff, excel_rows = load_attendance_data_from_api(selected_date_str, selected_date)
    
    tot_val = len(active_employees)
    pre_val = len(present_staff)
    lat_val = len(late_staff)
    chk_val = len(checkout_staff)
    abs_val = len(full_absent_staff)

    # Dynamic Spacious KPI Grid
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        if st.button(f"👥 الكل\n\n{tot_val}", key="kpi_all"):
            st.session_state["selected_view"] = "all"; st.rerun()
    with c2:
        if st.button(f"🟢 بالعمل\n\n{pre_val}", key="kpi_present"):
            st.session_state["selected_view"] = "present"; st.rerun()
    with c3:
        if st.button(f"⏰ متأخر\n\n{lat_val}", key="kpi_late"):
            st.session_state["selected_view"] = "late"; st.rerun()
    with c4:
        if st.button(f"🏁 انصراف\n\n{chk_val}", key="kpi_checkout"):
            st.session_state["selected_view"] = "checkout"; st.rerun()
    with c5:
        if st.button(f"❌ غياب\n\n{abs_val}", key="kpi_absent"):
            st.session_state["selected_view"] = "absent"; st.rerun()

    search_query = st.text_input("", placeholder=TEXT_CONFIG["search_placeholder"], label_visibility="collapsed").strip().lower()

    def match_search(code, name):
        if not search_query: return True
        return search_query in str(code).lower() or search_query in str(name).lower()

    excel_buffer = io.BytesIO()
    df_grid_data = pd.DataFrame(excel_rows)
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df_grid_data.to_excel(writer, sheet_name="Attendance Report", index=False, startrow=5)
        workbook = writer.book
        worksheet = writer.sheets["Attendance Report"]
        worksheet["A1"] = "Daily Attendance Report (Basic Report)"
        worksheet["A1"].font = Font(name="Arial", size=16, bold=True)
        worksheet.merge_cells("A1:H1")
        
        for col_cells in worksheet.columns:
            max_len = 0
            for cell in col_cells:
                if cell.row > 5 and cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            col_letter = get_column_letter(col_cells[0].column)
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    current_view = st.session_state["selected_view"]

    if current_view == "all":
        filtered = [f"<tr><td>{c}</td><td>{n}</td><td><span class='badge badge-present'>نشط</span></td></tr>" for c, n in active_employees.items() if match_search(c, n)]
        st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="3" class="table-main-title-header">{TEXT_CONFIG["header_all"].format(tot_val)}</th></tr><tr><th>الكود</th><th>اسم الموظف</th><th>الحالة</th></tr>{"".join(filtered)}</table>', unsafe_allow_html=True)
        
    elif current_view == "present":
        if present_staff:
            filtered = [f"<tr><td>{c}</td><td>{n}</td><td>{t}</td><td>{d}</td></tr>" for c, n, t, d in present_staff if match_search(c, n)]
            st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="4" class="table-main-title-header">{TEXT_CONFIG["header_present"].format(pre_val)}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الدخول</th><th>الجهاز</th></tr>{"".join(filtered)}</table>', unsafe_allow_html=True)
        else: st.info("لا يوجد موظفين متواجدين حالياً.")
        
    elif current_view == "late":
        if late_staff:
            filtered = [f"<tr><td>{c}</td><td>{n}</td><td>{t}</td><td><span class='badge badge-late'>متأخر</span></td></tr>" for c, n, t, d in late_staff if match_search(c, n)]
            st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="4" class="table-main-title-header">{TEXT_CONFIG["header_late"].format(lat_val)}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الدخول</th><th>الحالة</th></tr>{"".join(filtered)}</table>', unsafe_allow_html=True)
        else: st.success("🎉 لا يوجد متأخرين لهذا اليوم!")
        
    elif current_view == "checkout":
        if checkout_staff:
            filtered = [f"<tr><td>{c}</td><td>{n}</td><td>{t}</td><td>{d}</td></tr>" for c, n, t, d in checkout_staff if match_search(c, n)]
            st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="4" class="table-main-title-header">{TEXT_CONFIG["header_checkout"].format(chk_val)}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الانصراف</th><th>الجهاز</th></tr>{"".join(filtered)}</table>', unsafe_allow_html=True)
        else: st.info("لا توجد عمليات انصراف مسجلة حتى الآن.")
        
    elif current_view == "absent":
        if full_absent_staff:
            filtered = [f"<tr><td>{c}</td><td>{n}</td><td><span class='badge badge-absent'>غياب</span></td></tr>" for c, n in full_absent_staff if match_search(c, n)]
            st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="3" class="table-main-title-header">{TEXT_CONFIG["header_absent"].format(abs_val)}</th></tr><tr><th>الكود</th><th>اسم الموظف</th><th>الحالة</th></tr>{"".join(filtered)}</table>', unsafe_allow_html=True)
        else: st.success("🎉 لا يوجد غيابات لهذا اليوم!")

    st.markdown("---")
    st.download_button(label=TEXT_CONFIG["btn_download_excel"], data=excel_buffer.getvalue(), file_name=f"Daily_Attendance_Report_{selected_date_str}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

except Exception as e:
    st.error(TEXT_CONFIG["err_api"].format(str(e)))
