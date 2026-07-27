import streamlit as st
import requests
import unicodedata
import pandas as pd
from datetime import datetime
import zoneinfo

# ==========================================
# 0. RTL ARABIC TEXT & VISUAL CONFIG
# ==========================================
TEXT_CONFIG = {
    "page_title": "حضور وانصراف القصر الذهبي",
    "title_main": "✨ شركة القصر الذهبي ✨",
    "lbl_date": "📅 التاريخ الحالي في سوريا: **{}**  │  ⏰ الوقت الحالي: **{}**",
    "btn_refresh": "🔄 تحديث البيانات الحية الآن",
    "btn_download_excel": "📥 تحميل تقرير الحضور الشامل (CSV/Excel)",
    
    # رؤوس القوائم التفصيلية
    "header_late": "⏰ قائمة الموظفين المتأخرين اليوم ({})",
    "header_absent": "❌ قائمة الغيابات الكاملة اليوم ({})",
    "header_present": "🟢 قائمة الموظفين المتواجدون حالياً ({})",
    "header_checkout": "🏁 قائمة الموظفين المنصرفين اليوم ({})",
    "header_all": "👥 قائمة كافة موظفي الشركة النشطين ({})",
    
    # نصوص أسطر العرض
    "late_row": "🔸 **{}** (كود: {}) ── وقت الدخول: {}",
    "absent_row": "🔹 **{}** (كود: {})",
    "present_row": "🔸 **{}** (كود: {}) ── وقت الدخول: {}",
    "checkout_row": "✅ **{}** (كود: {}) ── وقت الانصراف: {}",
    "all_row": "👤 **{}** (كود: {})",
    
    "err_api": "خطأ في الاتصال بواجهة BioTime السحابية: {}"
}

st.set_page_config(page_title=TEXT_CONFIG["page_title"], page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; }
    .reportview-container .main .block-container { direction: RTL; text-align: right; }
    h1, h2, h3, h4, p, span, li, div { text-align: right !important; direction: RTL !important; line-height: 1.6 !important; }
    
    /* تنسيق أزرار بطاقات الحضور */
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
        margin-bottom: 5px !important;
    }
    div.stButton > button:hover {
        border-color: #cbd5e1 !important;
        background-color: #f8fafc !important;
    }
    
    /* بطاقات حالات الأجهزة المضافة */
    .device-box {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        padding: 10px 15px;
        margin-bottom: 10px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .status-online { color: #10b981; font-weight: bold; }
    .status-offline { color: #ef4444; font-weight: bold; }
    
    .list-wrapper-box {
        background-color: #f8fafc;
        border: 1px dashed #cbd5e1;
        border-radius: 6px;
        padding: 15px;
        margin-top: 10px;
        margin-bottom: 25px;
        direction: rtl;
    }
    </style>
""", unsafe_allow_html=True)

# استثناءات الإدارة والموظفين المستقيلين
EXCLUDED_MANAGEMENT_CODES = ("40", "10", "20")
EXCLUDED_RESIGNED_CODES = ("105", "112", "130") # 📝 اكتب هنا الأكواد الفردية للمستقيلين

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
        raise Exception("تفاصيل رمز المصادقة (Token) مفقودة أو غير صالحة.")
        
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    st.session_state["debug_logs"] = []
    
    # 1. جلب الموظفين الصافيين غير المستبعدين
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
        if code and code not in EXCLUDED_MANAGEMENT_CODES and code not in EXCLUDED_RESIGNED_CODES:
            first_name = emp.get("first_name", "") or ""
            last_name = emp.get("last_name", "") or ""
            full_name = f"{first_name} {last_name}".strip()
            active_employees[code] = clean_txt(full_name if full_name else f"User {code}")

    # 2. جلب حالة أجهزة البصمة المضافة حديثاً لمراقبة السيرفر
    device_url = f"{BASE_URL}/iclock/api/devices/?page_size=100"
    device_list = []
    try:
        dev_res = requests.get(device_url, headers=headers, timeout=10)
        if dev_res.status_code == 200:
            device_list = dev_res.json().get("data", [])
    except Exception as e:
        log_debug(f"Device Request Error: {str(e)}")

    # 3. جلب حركات البصمات اليومية
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

    for code, name in active_employees.items():
        if code in emp_punches and emp_punches[code]:
            user_punches = emp_punches[code]
            first_punch = user_punches
            last_punch = user_punches[-1]
            punch_count = len(user_punches)
            
            time_in_clean = first_punch.strftime('%I:%M %p')
            time_out_clean = last_punch.strftime('%I:%M %p')

            if first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15):
                late_staff.append((code, name, time_in_clean))

            if punch_count % 2 != 0:
                present_staff.append((code, name, time_in_clean))
            else:
                checkout_staff.append((code, name, time_out_clean))
        else:
            full_absent_staff.append((code, name))

    # FIXED: فحص وترتيب آمن يمنع خطأ الـ tuple object has no attribute 'isdigit' نهائياً
    full_absent_staff.sort(key=lambda x: int(x[0]) if str(x[0]).isdigit() else str(x[0]))
    present_staff.sort(key=lambda x: int(x[0]) if str(x[0]).isdigit() else str(x[0]))
    late_staff.sort(key=lambda x: int(x[0]) if str(x[0]).isdigit() else str(x[0]))
    checkout_staff.sort(key=lambda x: int(x[0]) if str(x[0]).isdigit() else str(x[0]))
    
    return active_employees, present_staff, late_staff, full_absent_staff, checkout_staff, device_list
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
    active_employees, present_staff, late_staff, full_absent_staff, checkout_staff, device_list = load_attendance_data_from_api(selected_date_str)
    
    report_rows = []
    for c, n, t in present_staff:
        report_rows.append({"كود الموظف": c, "الاسم": n, "وقت الدخول": t, "وقت الانصراف": "متواجد حالياً", "الحالة اليومية": "متواجد في العمل"})
    for c, n, t in late_staff:
        report_rows.append({"كود الموظف": c, "الاسم": n, "وقت الدخول": t, "وقت الانصراف": "-", "الحالة اليومية": "متأخر صباحاً"})
    for c, n, t in checkout_staff:
        report_rows.append({"كود الموظف": c, "الاسم": n, "وقت الدخول": "-", "وقت الانصراف": t, "الحالة اليومية": "سجل انصراف"})
    for c, n in full_absent_staff:
        report_rows.append({"كود الموظف": c, "الاسم": n, "وقت الدخول": "-", "وقت الانصراف": "-", "الحالة اليومية": "غياب كامل"})
        
    df_report = pd.DataFrame(report_rows)
    csv_data = df_report.to_csv(index=False).encode('utf-8-sig')
    
    with btn_col2:
        st.download_button(
            label=TEXT_CONFIG["btn_download_excel"],
            data=csv_data,
            file_name=f"تقرير_حضور_القصر_الذهبي_{selected_date_str}.csv",
            mime="text/csv",
            use_container_width=True
        )

    # حساب المجاميع الصافية
    total_emp = len(active_employees)
    p_count = len(present_staff)
    l_count = len(late_staff)
    c_count = len(checkout_staff)
    a_count = len(full_absent_staff)
    
    st.write("### 📊 اضغط على أي بطاقة لعرض أسماء الموظفين")
    
    if st.button(f"👥 إجمالي عدد موظفي الشركة النشطين ── {total_emp}"):
        st.session_state["selected_view"] = "all"
        
    if st.button(f"🟢 الموظفون المتواجدون حالياً في العمل ── {p_count}"):
        st.session_state["selected_view"] = "present"
        
    if st.button(f"⏰ الموظفون المتأخرون اليوم ── {l_count}"):
        st.session_state["selected_view"] = "late"
        
    if st.button(f"✅ الموظفون الذين غادروا وانصرفوا ── {c_count}"):
        st.session_state["selected_view"] = "checkout"
        
    if st.button(f"❌ الموظفون الغائبون بالكامل اليوم ── {a_count}"):
        st.session_state["selected_view"] = "absent"

    # 📡 استعراض حالة أجهزة البصمة المضافة بناءً على طلبك
    st.write("### 📡 حالة الاتصال الحية لأجهزة البصمة")
    if device_list:
        dev_col1, dev_col2 = st.columns(2)
        for idx, dev in enumerate(device_list):
            dev_name = dev.get("alias", "جهاز غير مسمى")
            is_online = dev.get("state", False) or (str(dev.get("status", "")).lower() == "online")
            
            status_html = '<span class="status-online">🟢 متصل الآن (Online)</span>' if is_online else '<span class="status-offline">🔴 منقطع (Offline)</span>'
            card_html = f'<div class="device-box"><span>📟 {dev_name}</span>{status_html}</div>'
            
            if idx % 2 == 0:
                with dev_col1: st.markdown(card_html, unsafe_allow_html=True)
            else:
                with dev_col2: st.markdown(card_html, unsafe_allow_html=True)
    else:
        st.info("لا توجد أجهزة بصمة مسجلة أو مرئية حالياً في الحساب.")

    # ------------------------------------------
    # مساحة استعراض القوائم التفاعلية المحدثة والمحمية
    # ------------------------------------------
    current_view = st.session_state["selected_view"]
    
    if current_view == "all":
        st.write(f"### {TEXT_CONFIG['header_all'].format(total_emp)}")
        st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
        for code, name in sorted(active_employees.items(), key=lambda x: int(x[0]) if x[0].isdigit() else x[0]):
            st.markdown(TEXT_CONFIG["all_row"].format(name, code))
        st.markdown('</div>', unsafe_allow_html=True)
        
    elif current_view == "present":
        st.write(f"### {TEXT_CONFIG['header_present'].format(p_count)}")
        if present_staff:
            st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
            for code, name, time_in in present_staff:
                st.markdown(TEXT_CONFIG["present_row"].format(name, code, time_in))
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("لا يوجد موظفين متواجدين حالياً.")
            
    elif current_view == "late":
        st.write(f"### {TEXT_CONFIG['header_late'].format(l_count)}")
        if late_staff:
            st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
            for code, name, time_in in late_staff:
                st.markdown(TEXT_CONFIG["late_row"].format(name, code, time_in))
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.success("🎉 لا يوجد متأخرين اليوم!")
            
    elif current_view == "checkout":
        st.write(f"### {TEXT_CONFIG['header_checkout'].format(c_count)}")
        if checkout_staff:
            st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
            for code, name, time_out in checkout_staff:
                st.markdown(TEXT_CONFIG["checkout_row"].format(name, code, time_out))
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("لا توجد عمليات انصراف مسجلة حتى الآن.")
            
    elif current_view == "absent":
        st.write(f"### {TEXT_CONFIG['header_absent'].format(a_count)}")
        if full_absent_staff:
            st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
            for code, name in full_absent_staff:
                st.markdown(TEXT_CONFIG["absent_row"].format(name, code))
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.success("🎉 لا يوجد غيابات اليوم!")

except Exception as e:
    st.error(TEXT_CONFIG["err_api"].format(str(e)))
    if st.checkbox("عرض سجلات الأخطاء البرمجية (Debug Logs)"):
        for log in st.session_state.get("debug_logs", []):
            st.text(log)
