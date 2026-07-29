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
    "btn_refresh": "🔄 تحديث البيانات الحية",
    "lbl_pick_date": "📅 اختر تاريخ عرض التقرير:",
    "btn_download_excel": "📥 تحميل تقرير Excel النمطي",
    "search_placeholder": "🔍 ابحث باسم الموظف أو رقم الكود...",
    
    "header_all": "👥 كافة موظفي الشركة النشطين ({})",
    "header_present": "🟢 الموظفون المتواجدون حالياً ({})",
    "header_late": "⏰ الموظفون المتأخرين اليوم ({})",
    "header_checkout": "🏁 الموظفون المنصرفون اليوم ({})",
    "header_absent": "❌ قائمة الغيابات الكاملة اليوم ({})",
    "err_api": "خطأ في الاتصال بواجهة BioTime السحابية: {}"
}

st.set_page_config(page_title=TEXT_CONFIG["page_title"], page_icon="📊", layout="wide", initial_sidebar_state="collapsed")

# 📱 ULTIMATE NATIVE MOBILE APP SHELL OVERRIDE (HTML/CSS DECK)
st.markdown("""
    <style>
    /* 1. Eliminate Streamlit Desktop Browser Framework */
    header[data-testid="stHeader"] { display: none !important; }
    footer { display: none !important; }
    div[data-testid="stDecoration"] { display: none !important; }
    .stApp { direction: rtl; background-color: #f8fafc; }
    
    /* 2. Absolute Screen Margin Lock for Mobile Viewports */
    .reportview-container .main .block-container {
        padding-top: 15px !important;
        padding-bottom: 90px !important; /* Spacing for the sticky bottom navbar */
        padding-left: 12px !important;
        padding-right: 12px !important;
        max-width: 100% !important;
    }
    
    /* 3. Mobile Stat Summary Ring Grid */
    .kpi-container {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 8px;
        margin: 15px 0;
    }
    .kpi-card {
        background: #ffffff;
        border-radius: 10px;
        padding: 10px 4px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        border: 1px solid #e2e8f0;
    }
    .kpi-val { font-size: 18px; font-weight: 800; margin-bottom: 2px; }
    .kpi-lbl { font-size: 11px; color: #64748b; font-weight: bold; }
    
    /* 4. Native App Native Sticky Bottom Navigation Bar Matrix */
    .bottom-navbar {
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        height: 65px;
        background-color: #ffffff;
        border-top: 1px solid #e2e8f0;
        display: flex;
        justify-content: space-around;
        align-items: center;
        z-index: 999999;
        padding-bottom: constant(safe-area-inset-bottom); /* iOS notch optimization */
        padding-bottom: env(safe-area-inset-bottom);
        box-shadow: 0 -4px 10px rgba(0,0,0,0.04);
    }
    
    /* 5. Mobile Status Badges Layout */
    .badge {
        padding: 4px 8px;
        border-radius: 6px;
        font-size: 11px;
        font-weight: bold;
    }
    .badge-present { background-color: #dcfce7; color: #166534; }
    .badge-late { background-color: #fef3c7; color: #9a3412; }
    .badge-absent { background-color: #fee2e2; color: #991b1b; }
    
    /* 6. Native App Table Framework Rows */
    .responsive-grid-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 10px;
        font-size: 13px;
        background-color: #ffffff;
        border-radius: 10px;
        overflow: hidden;
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
        background-color: #f1f5f9;
        color: #475569;
        padding: 10px;
        font-weight: 700;
        border-bottom: 1px solid #e2e8f0;
    }
    .responsive-grid-table td { padding: 10px; border-bottom: 1px solid #f8fafc; color: #334155; }
    
    .miracle-banner {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        color: #ffffff !important;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        margin: 15px 0;
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
    excel_rows.sort(key=lambda row_item: int(row_item["Employee ID"]) if row_item["Employee ID"].isdigit() else 999)
    
    return active_employees, present_staff, late_staff, full_absent_staff, checkout_staff, excel_rows
# ==========================================
# 3. INTERFACE RENDERING & CONTROLS MATRIX
# ==========================================
now_syria = datetime.now(SYRIA_TZ)
time_str = now_syria.strftime('%I:%M:%S %p')

# 📱 COMPACT MOBILE HEADER SECTION
col_left, col_right = st.columns([3, 1])
with col_left:
    selected_date = st.date_input("", value=now_syria.date(), label_visibility="collapsed")
    selected_date_str = selected_date.strftime('%Y-%m-%d')
with col_right:
    if st.button("🔄", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

formatted_excel_date = selected_date.strftime('%B %d %Y')
generated_on_timestamp = now_syria.strftime('%a %b %d %Y %H:%M:%S')

try:
    active_employees, present_staff, late_staff, full_absent_staff, checkout_staff, excel_rows = load_attendance_data_from_api(selected_date_str, selected_date)
    is_today = (selected_date == now_syria.date())
    
    # 📊 NATIVE MOBILE MINI KPI BLOCK ROW
    tot_val = len(active_employees)
    pre_val = len(present_staff) if is_today else 0
    lat_val = len(late_staff) if is_today else 0
    chk_val = len(checkout_staff) if is_today else 0
    abs_val = len(full_absent_staff) if is_today else 0

    st.markdown(f"""
        <div class="kpi-container">
            <div class="kpi-card" style="border-bottom: 3px solid #1e3a8a;"><div class="kpi-val" style="color: #1e3a8a;">{tot_val}</div><div class="kpi-lbl">الكل</div></div>
            <div class="kpi-card" style="border-bottom: 3px solid #16a34a;"><div class="kpi-val" style="color: #16a34a;">{pre_val}</div><div class="kpi-lbl">بالعمل</div></div>
            <div class="kpi-card" style="border-bottom: 3px solid #ea580c;"><div class="kpi-val" style="color: #ea580c;">{lat_val}</div><div class="kpi-lbl">متأخر</div></div>
            <div class="kpi-card" style="border-bottom: 3px solid #dc2626;"><div class="kpi-val" style="color: #dc2626;">{abs_val}</div><div class="kpi-lbl">غياب</div></div>
        </div>
    """, unsafe_allow_html=True)

    # 🔍 IN-LINE LIVE FILTER INPUT FIELD 
    search_query = st.text_input("", placeholder=TEXT_CONFIG["search_placeholder"], label_visibility="collapsed").strip().lower()

    def match_search(code, name):
        if not search_query: return True
        return search_query in str(code).lower() or search_query in str(name).lower()

    def show_miracle_message():
        st.markdown(f"""
            <div class="miracle-banner">
                <div class="miracle-title">✨ هذا اليوم مؤرشف بالكامل ✨</div>
                <div class="miracle-text">تم قفل وحفظ سجلات يوم <b>{selected_date_str}</b> بأمان في قاعدة البيانات السحابية.</div>
            </div>
        """, unsafe_allow_html=True)

    # Excel builder engine background task 
    excel_buffer = io.BytesIO()
    df_grid_data = pd.DataFrame(excel_rows)
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df_grid_data.to_excel(writer, sheet_name="Attendance Report", index=False, startrow=5)
        workbook = writer.book
        worksheet = writer.sheets["Attendance Report"]
        worksheet["A1"] = "Daily Attendance Report(Basic Report)"
        worksheet["A1"].font = Font(name="Arial", size=16, bold=True)
        worksheet.merge_cells("A1:H1")
        for col_cells in worksheet.columns:
            max_len = max([len(str(cell.value or '')) for cell in col_cells if cell.row > 5] +)
            worksheet.column_dimensions[get_column_letter(col_cells.column)].width = max_len + 4

    current_view = st.session_state["selected_view"]

    # 🗂️ RENDER SELECTED CELL CONTENT DATA 
    if current_view == "all":
        if is_today:
            filtered = [f"<tr><td>{c}</td><td>{n}</td><td><span class='badge badge-present'>نشط</span></td></tr>" for c, n in active_employees.items() if match_search(c, n)]
            st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="3" class="table-main-title-header">{TEXT_CONFIG["header_all"].format(tot_val)}</th></tr><tr><th>الكود</th><th>اسم الموظف</th><th>الحالة</th></tr>{"".join(filtered)}</table>', unsafe_allow_html=True)
        else: show_miracle_message()
        
    elif current_view == "present":
        if is_today:
            if present_staff:
                filtered = [f"<tr><td>{c}</td><td>{n}</td><td>{t}</td><td>{d}</td></tr>" for c, n, t, d in present_staff if match_search(c, n)]
                st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="4" class="table-main-title-header">{TEXT_CONFIG["header_present"].format(pre_val)}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الدخول</th><th>الجهاز</th></tr>{"".join(filtered)}</table>', unsafe_allow_html=True)
            else: st.info("لا يوجد موظفين متواجدين حالياً.")
        else: show_miracle_message()
        
    elif current_view == "late":
        if is_today:
            if late_staff:
                filtered = [f"<tr><td>{c}</td><td>{n}</td><td>{t}</td><td><span class='badge badge-late'>متأخر</span></td></tr>" for c, n, t, d in late_staff if match_search(c, n)]
                st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="4" class="table-main-title-header">{TEXT_CONFIG["header_late"].format(lat_val)}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الدخول</th><th>الحالة</th></tr>{"".join(filtered)}</table>', unsafe_allow_html=True)
            else: st.success("🎉 لا يوجد متأخرين اليوم!")
        else: show_miracle_message()
        
    elif current_view == "checkout":
        if is_today:
            if checkout_staff:
                filtered = [f"<tr><td>{c}</td><td>{n}</td><td>{t}</td><td>{d}</td></tr>" for c, n, t, d in checkout_staff if match_search(c, n)]
                st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="4" class="table-main-title-header">{TEXT_CONFIG["header_checkout"].format(chk_val)}</th></tr><tr><th>الكود</th><th>الاسم</th><th>الانصراف</th><th>الجهاز</th></tr>{"".join(filtered)}</table>', unsafe_allow_html=True)
            else: st.info("لا توجد عمليات انصراف مسجلة.")
        else: show_miracle_message()
        
    elif current_view == "absent":
        if is_today:
            if full_absent_staff:
                filtered = [f"<tr><td>{c}</td><td>{n}</td><td><span class='badge badge-absent'>غياب</span></td></tr>" for c, n in full_absent_staff if match_search(c, n)]
                st.markdown(f'<table class="responsive-grid-table"><tr><th colspan="3" class="table-main-title-header">{TEXT_CONFIG["header_absent"].format(abs_val)}</th></tr><tr><th>الكود</th><th>اسم الموظف</th><th>الحالة</th></tr>{"".join(filtered)}</table>', unsafe_allow_html=True)
            else: st.success("🎉 لا يوجد غيابات اليوم!")
        else: show_miracle_message()

    # 📱 STICKY BUTTON TAB NAVIGATION BAR FOR ACTIVE TOUCH CONTROL
    col_t1, col_v2, col_v3, col_v4, col_v5 = st.columns(5)
    st.markdown("""
        <div class="bottom-navbar">
            <a href="javascript:void(0);" onclick="window.parent.postMessage({type: 'streamlit:set_component_value', value: 'all'}, '*')" style="text-decoration:none; text-align:center; color:#64748b;"><div>👤</div><div style="font-size:10px;">الكل</div></a>
            <a href="javascript:void(0);" onclick="window.parent.postMessage({type: 'streamlit:set_component_value', value: 'present'}, '*')" style="text-decoration:none; text-align:center; color:#16a34a;"><div>🟢</div><div style="font-size:10px;">بالعمل</div></a>
            <a href="javascript:void(0);" onclick="window.parent.postMessage({type: 'streamlit:set_component_value', value: 'late'}, '*')" style="text-decoration:none; text-align:center; color:#ea580c;"><div>⏰</div><div style="font-size:10px;">متأخر</div></a>
            <a href="javascript:void(0);" onclick="window.parent.postMessage({type: 'streamlit:set_component_value', value: 'checkout'}, '*')" style="text-decoration:none; text-align:center; color:#2563eb;"><div>🏁</div><div style="font-size:10px;">انصراف</div></a>
            <a href="javascript:void(0);" onclick="window.parent.postMessage({type: 'streamlit:set_component_value', value: 'absent'}, '*')" style="text-decoration:none; text-align:center; color:#dc2626;"><div>❌</div><div style="font-size:10px;">غياب</div></a>
        </div>
    """, unsafe_allow_html=True)
    
    # Standard fallback interaction grid layout for stable device handling
    st.markdown("---")
    st.write("⚙️ **لوحة التحكم السريعة:**")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        if st.button("👤 الكل"): st.session_state["selected_view"] = "all"; st.rerun()
    with c2:
        if st.button("🟢 بالعمل"): st.session_state["selected_view"] = "present"; st.rerun()
    with c3:
        if st.button("⏰ متأخر"): st.session_state["selected_view"] = "late"; st.rerun()
    with c4:
        if st.button("🏁 انصراف"): st.session_state["selected_view"] = "checkout"; st.rerun()
    with c5:
        if st.button("❌ غياب"): st.session_state["selected_view"] = "absent"; st.rerun()
        
    st.download_button(label=TEXT_CONFIG["btn_download_excel"], data=excel_buffer.getvalue(), file_name=f"Daily_Attendance_Report_{selected_date_str}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

except Exception as e:
    st.error(TEXT_CONFIG["err_api"].format(str(e)))
