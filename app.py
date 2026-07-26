import streamlit as st
import requests
import unicodedata
import pandas as pd
import io
from datetime import datetime
import zoneinfo

TEXT_CONFIG = {
    "page_title": "حضور وانصراف القصر الذهبي",
    "title_main": "✨ شركة القصر الذهبي ✨",
    "lbl_date": "📅 التاريخ المختار للتقرير: **{}**  │  ⏰ الوقت الحالي في سوريا: **{}**",
    "btn_refresh": "🔄 تحديث البيانات الحية الآن",
    "lbl_pick_date": "📅 اختر تاريخ عرض التقرير:",
    "btn_download_excel": "📥 تحميل تقرير الحضور كملف Excel",
    "header_late": "⏰ المتأخرون اليوم ({}) – دخول بعد 09:15 صباحاً",
    "header_absent": "❌ غائبون أو نسوا تسجيل الحضور ({})",
    "header_present": "🟢 الموظفون المتواجدون حالياً في العمل ({})",
    "header_early_leave": "⚠️ غادروا العمل مبكراً اليوم ({}) – خروج قبل 04:00 مساءً",
    "header_checkout": "🏁 الموظفون الذين غادروا بانتظام ({})",
    "late_row": "🔸 **{}** (كود: {}) ── وقت الدخول: {}",
    "absent_row": "🔹 **{}** (كود: {})",
    "present_row": "🔸 **{}** (كود: {}) ── وقت الدخول: {}",
    "early_leave_row": "⚠️ **{}** (كود: {}) ── الدخول: {} ── الخروج المبكر: {}",
    "checkout_row": "✅ **{}** (كود: {}) ── وقت الانصراف: {}",
    "success_no_late": "🎉 لا يوجد متأخرين اليوم!",
    "success_no_absent": "🎉 لا يوجد غيابات اليوم!",
    "info_no_present": "لا يوجد موظفين متواجدين حالياً في المنشأة.",
    "success_no_early": "🎉 لا يوجد حالات خروج مبكر اليوم!",
    "info_no_checkout": "لا توجد عمليات انصراف مسجلة حتى الآن.",
    "err_api": "خطأ في الاتصال بواجهة BioTime السحابية: {}"
}

st.set_page_config(page_title=TEXT_CONFIG["page_title"], page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; }
    .reportview-container .main .block-container { direction: RTL; text-align: right; }
    h1, h2, h3, h4, p, span, li, div { text-align: right !important; direction: RTL !important; line-height: 1.6 !important; }
    .metric-card {
        background-color: #ffffff;
        border-radius: 8px;
        padding: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.06);
        border: 1px solid #eef2f5;
        display: flex;
        align-items: center;
        margin-bottom: 12px;
    }
    .metric-icon {
        width: 40px;
        height: 40px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 20px;
        margin-left: 12px;
    }
    .icon-total { background-color: #e2e8f0; color: #475569; }
    .icon-present { background-color: #e8f5e9; color: #4caf50; }
    .icon-absent { background-color: #ffebee; color: #f44336; }
    .icon-late { background-color: #fff8e1; color: #ffc107; }
    .icon-early { background-color: #fff3e0; color: #ff9800; }
    .icon-checkout { background-color: #e0f7fa; color: #00bcd4; }
    .metric-info { display: flex; flex-direction: column; flex-grow: 1; }
    .metric-title { font-size: 13px; color: #64748b; font-weight: 500; }
    .metric-value { font-size: 22px; font-weight: bold; color: #1e293b; }
    .list-wrapper-box {
        background-color: #f8fafc;
        border: 1px dashed #cbd5e1;
        border-radius: 6px;
        padding: 12px;
        margin-bottom: 15px;
        direction: rtl;
    }
    </style>
""", unsafe_allow_html=True)

EXCLUDED_MANAGEMENT_CODES = ("40", "10", "20")
SYRIA_TZ = zoneinfo.ZoneInfo("Asia/Damascus")

BASE_URL = st.secrets["biotime"]["base_url"].rstrip('/')
TOKEN_URL = st.secrets["biotime"]["token_url"]
EMAIL = st.secrets["biotime"]["email"]
PASSWORD = st.secrets["biotime"]["password"]
COMPANY = st.secrets["biotime"]["company"]

if "debug_logs" not in st.session_state:
    st.session_state["debug_logs"] = []

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
        # FIXED: Added array parameters [200, 201] inside evaluation loop
        if response.status_code in:
            return response.json().get("token")
        return None
    except Exception:
        return None

def load_attendance_data_from_api(selected_date_str):
    token = get_auth_token()
    if not token:
        raise Exception("تفاصيل رمز المصادقة (Token) مفقودة أو غير صالحة.")
        
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    st.session_state["debug_logs"] = []
    
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
    early_leave_staff = []  
    checkout_staff = []     

    for code, name in active_employees.items():
        if code in emp_punches and emp_punches[code]:
            user_punches = emp_punches[code]
            first_punch = user_punches[0]
            last_punch = user_punches[-1]
            punch_count = len(user_punches)
            
            time_in_clean = first_punch.strftime('%I:%M %p')
            time_out_clean = last_punch.strftime('%I:%M %p')

            if first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15):
                late_staff.append((code, name, time_in_clean))

            if punch_count % 2 != 0:
                present_staff.append((code, name, time_in_clean))
            else:
                if last_punch.hour < 16:
                    early_leave_staff.append((code, name, time_in_clean, time_out_clean))
                else:
                    checkout_staff.append((code, name, time_out_clean))
        else:
            full_absent_staff.append((code, name))

    full_absent_staff.sort(key=lambda x: int(x) if x.isdigit() else x)
    present_staff.sort(key=lambda x: int(x) if x.isdigit() else x)
    late_staff.sort(key=lambda x: int(x) if x.isdigit() else x)
    early_leave_staff.sort(key=lambda x: int(x) if x.isdigit() else x)
    checkout_staff.sort(key=lambda x: int(x) if x.isdigit() else x)
    
    return active_employees, present_staff, late_staff, full_absent_staff, early_leave_staff, checkout_staff
now_syria = datetime.now(SYRIA_TZ)
time_str = now_syria.strftime('%I:%M:%S %p')

selected_date = st.date_input(TEXT_CONFIG["lbl_pick_date"], value=now_syria.date())
selected_date_str = selected_date.strftime('%Y-%m-%d')

st.title(TEXT_CONFIG["title_main"])
st.markdown(TEXT_CONFIG["lbl_date"].format(selected_date_str, time_str))

btn_col1, btn_col2 = st.columns(2)
with btn_col1:
    if st.button(TEXT_CONFIG["btn_refresh"], use_container_width=True):
        st.cache_data.clear()
        st.rerun()

try:
    active_employees, present_staff, late_staff, full_absent_staff, early_leave_staff, checkout_staff = load_attendance_data_from_api(selected_date_str)
    
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        pd.DataFrame([{"كود الموظف": c, "الاسم": n, "وقت الدخول": t} for c, n, t in present_staff]).to_excel(writer, sheet_name="متواجدون حاليا", index=False)
        pd.DataFrame([{"كود الموظف": c, "الاسم": n, "وقت الدخول": t} for c, n, t in late_staff]).to_excel(writer, sheet_name="المتأخرون", index=False)
        pd.DataFrame([{"كود الموظف": c, "الاسم": n, "وقت الدخول": t_in, "وقت الانصراف المبكر": t_out} for c, n, t_in, t_out in early_leave_staff]).to_excel(writer, sheet_name="خروج مبكر", index=False)
        pd.DataFrame([{"كود الموظف": c, "الاسم": n, "وقت الانصراف": t} for c, n, t in checkout_staff]).to_excel(writer, sheet_name="انصراف نظامي", index=False)
        pd.DataFrame([{"كود الموظف": c, "الاسم": n} for c, n in full_absent_staff]).to_excel(writer, sheet_name="غياب كامل", index=False)
    
    with btn_col2:
        st.download_button(
            label=TEXT_CONFIG["btn_download_excel"],
            data=excel_buffer.getvalue(),
            file_name=f"حضور_القصر_الذهبي_{selected_date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.write("### 📊 إحصائيات الحالة العامة للموظفين")
    total_emp = len(active_employees)
    p_count = len(present_staff)
    a_count = len(full_absent_staff)
    l_count = len(late_staff)
    e_count = len(early_leave_staff)
    c_count = len(checkout_staff)
    
    m_col1, m_col2, m_col3, m_col4, m_col5, m_col6 = st.columns(6)
    with m_col1:
        st.markdown(f'<div class="metric-card"><div class="metric-icon icon-total">👥</div><div class="metric-info"><span class="metric-title">إجمالي العدد</span><span class="metric-value">{total_emp}</span></div></div>', unsafe_allow_html=True)
    with m_col2:
        st.markdown(f'<div class="metric-card"><div class="metric-icon icon-present">🟢</div><div class="metric-info"><span class="metric-title">المتواجدون</span><span class="metric-value">{p_count}</span></div></div>', unsafe_allow_html=True)
    with m_col3:
        st.markdown(f'<div class="metric-card"><div class="metric-icon icon-late">⏰</div><div class="metric-info"><span class="metric-title">المتأخرين</span><span class="metric-value">{l_count}</span></div></div>', unsafe_allow_html=True)
    with m_col4:
        st.markdown(f'<div class="metric-card"><div class="metric-icon icon-early">⚠️</div><div class="metric-info"><span class="metric-title">خروج مبكر</span><span class="metric-value">{e_count}</span></div></div>', unsafe_allow_html=True)
    with m_col5:
        st.markdown(f'<div class="metric-card"><div class="metric-icon icon-checkout">✅</div><div class="metric-info"><span class="metric-title">انصراف نظامي</span><span class="metric-value">{c_count}</span></div></div>', unsafe_allow_html=True)
    with m_col6:
        st.markdown(f'<div class="metric-card"><div class="metric-icon icon-absent">❌</div><div class="metric-info"><span class="metric-title">غياب كامل</span><span class="metric-value">{a_count}</span></div></div>', unsafe_allow_html=True)

    st.write("### 🔍 القوائم التفصيلية")
    
    with st.expander(TEXT_CONFIG["header_late"].format(l_count), expanded=True):
        if late_staff:
            st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
            for code, name, time_in in late_staff:
                st.markdown(TEXT_CONFIG["late_row"].format(name, code, time_in))
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.success(TEXT_CONFIG["success_no_late"])

    with st.expander(TEXT_CONFIG["header_present"].format(p_count), expanded=False):
        if present_staff:
            st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
            for code, name, time_in in present_staff:
                st.markdown(TEXT_CONFIG["present_row"].format(name, code, time_in))
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info(TEXT_CONFIG["info_no_present"])

    with st.expander(TEXT_CONFIG["header_early_leave"].format(e_count), expanded=False):
        if early_leave_staff:
            st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
            for code, name, time_in, time_out in early_leave_staff:
                st.markdown(TEXT_CONFIG["early_leave_row"].format(name, code, time_in, time_out))
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.success(TEXT_CONFIG["success_no_early"])

    with st.expander(TEXT_CONFIG["header_checkout"].format(c_count), expanded=False):
        if checkout_staff:
            st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
            for code, name, time_out in checkout_staff:
                st.markdown(TEXT_CONFIG["checkout_row"].format(name, code, time_out))
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info(TEXT_CONFIG["info_no_checkout"])

    with st.expander(TEXT_CONFIG["header_absent"].format(a_count), expanded=False):
        if full_absent_staff:
            st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
            for code, name in full_absent_staff:
                st.markdown(TEXT_CONFIG["absent_row"].format(name, code))
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.success(TEXT_CONFIG["success_no_absent"])

except Exception as e:
    st.error(TEXT_CONFIG["err_api"].format(str(e)))
    if st.checkbox("عرض سجلات الأخطاء البرمجية (Debug Logs)"):
        for log in st.session_state.get("debug_logs", []):
            st.text(log)
