import streamlit as st
import psycopg2
import pandas as pd
import unicodedata
from datetime import datetime
import zoneinfo
import urllib.parse
# ==========================================
# 0. CONFIGURATION DICTIONARY
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
    "title_sub": "لوحة تحكم إدارة الحضور والغياب",
    "lbl_date": "📅 التاريخ: **{}**  │  ⏰ الوقت الحالي في سوريا: **{}**",
    "btn_refresh": "🔄 تحديث البيانات الحية الآن",
    "status_online": "🟢 متصل",
    "status_offline": "🔴 غير متصل",
    "header_devices": "### 📡 حالة اتصال أجهزة البصمة الحالية:",
    "warn_no_devices": "⚠️ لا توجد أجهزة مضافة أو تعذر تحميل البيانات.",
    "header_late": "⏰ المتأخرون اليوم ({}) – دخول بعد 09:15 صباحاً",
    "late_row": "🔸 **{}** (كود: {}) ── وقت الدخول: {}",
    "success_no_late": "🎉 لا يوجد متأخرين اليوم!",
    "header_absent": "❌ غائبون أو نسوا تسجيل الحضور ({})",
    "absent_row": "🔹 **{}** (كود: {})",
    "sms_morning": "مرحباً {}، يظهر نظامنا أنك لم تقم بتسجيل الدخول اليوم. يرجى بصمة الدخول فوراً.",
    "btn_sms_morning": "💬 تذكير الدخول",
    "caption_no_phone": "🚫 لا يوجد رقم",
    "success_no_absent": "🎉 لا يوجد غيابات اليوم!",
    "header_present": "🟢 الموظفون المتواجدون حالياً في العمل ({})",
    "present_row": "🔸 **{}** (كود: {}) ── وقت الدخول: {}",
    "sms_evening": "مرحباً {}، لقد نسيت تسجيل الخروج اليوم. يرجى تذكر تبصيم الخروج قبل مغادرة العمل.",
    "btn_sms_evening": "💬 تذكير الخروج",
    "caption_locked_evening": "🔒 يفتح 06:45 مساءً",
    "info_no_present": "لا يوجد موظفين منتظمين متواجدين حالياً.",
    "err_db": "خطأ في الاتصال بقاعدة البيانات السحابية: {}",
    "header_live_log": "### 🔔 سجل البصمات الفوري (بث حي مباشر):",
    "live_log_row": "⚡ البصمة الأخيرة: قام **{}** (كود: {}) بالتبصيم الآن عند الساعة **{}**",
    "caption_no_logs": "⏳ بانتظار تسجيل أولى بصمات الموظفين اليوم..."
}

# ==========================================
# 1. INITIAL SYSTEM & WINDOW CONFIGURATION
# ==========================================
st.set_page_config(page_title="حضور القصر الذهبي", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .reportview-container .main .block-container { direction: RTL; text-align: right; }
    h1, h2, h3, h4, p, span, li, div { text-align: right !important; direction: RTL !important; line-height: 1.6 !important; }
    </style>
""", unsafe_allow_html=True)

# System Constants
EXCLUDED_MANAGEMENT_CODES = ("40", "10", "20")
mgmt_codes_str = ",".join(f"'{code}'" for code in EXCLUDED_MANAGEMENT_CODES)
DATABASE_URL = st.secrets["NEON_DATABASE_URL"]

# Explicitly lock the system clock to Syrian time boundaries
SYRIA_TZ = zoneinfo.ZoneInfo("Asia/Damascus")


# ==========================================
# 2. HELPER FUNCTIONS & LIVE DATA SERVICES
# ==========================================
def clean_phone(raw_phone):
    if not raw_phone: return ""
    clean_raw = unicodedata.normalize('NFKC', str(raw_phone)).encode('ascii', 'ignore').decode('ascii')
    phone = clean_raw.strip().replace(" ", "").replace("-", "").replace("+", "").lstrip("0")
    return f"963{phone}" if phone.startswith('9') and len(phone) == 9 else (f"963{phone[1:]}" if phone.startswith('09') else phone)

def clean_txt(raw_text):
    if not raw_text: return ""
    return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip())

def load_device_statuses():
    conn = psycopg2.connect(DATABASE_URL, sslmode='require')
    cursor = conn.cursor()
    device_metrics = []
    try:
        query = "SELECT alias, is_online, sn FROM iclock_terminal;"
        cursor.execute(query)
        rows = cursor.fetchall()
        for row in rows:
            alias, is_online, sn = row
            status_tag = "🟢 متصل" if is_online and (str(is_online).strip().lower() in ('true', '1', 't', 'y', 'yes')) else "🔴 غير متصل"
            device_metrics.append((clean_txt(alias), status_tag, sn))
    except Exception:
        conn.rollback()
        try:
            query = "SELECT alias, last_activity, sn FROM iclock_terminal;"
            cursor.execute(query)
            rows = cursor.fetchall()
            timestamps = [r for r in rows if r and r]
            latest_system_ping = max(timestamps) if timestamps else None
            for row in rows:
                alias, last_act, sn = row
                if last_act and latest_system_ping:
                    seconds_elapsed = (latest_system_ping.replace(tzinfo=None) - last_act.replace(tzinfo=None)).total_seconds()
                    status_tag = "🟢 متصل" if seconds_elapsed < 600 else "🔴 غير متصل"
                else:
                    status_tag = "🔴 غير متصل"
                device_metrics.append((clean_txt(alias), status_tag, sn))
        except Exception:
            pass
    finally:
        cursor.close()
        conn.close()
    return device_metrics

def load_live_punch_notifications(today_str):
    conn = psycopg2.connect(DATABASE_URL, sslmode='require')
    cursor = conn.cursor()
    recent_punches = []
    try:
        query = f"""
            SELECT e.emp_code, e.first_name, (t.punch_time AT TIME ZONE 'GMT-3') as p_time
            FROM iclock_transaction t
            JOIN personnel_employee e ON t.emp_id = e.id
            WHERE (t.punch_time AT TIME ZONE 'GMT-3')::date = '{today_str}'
            ORDER BY t.punch_time DESC LIMIT 5;
        """
        cursor.execute(query)
        rows = cursor.fetchall()
        for row in rows:
            emp_code, name, p_time = row
            recent_punches.append((clean_txt(emp_code), clean_txt(name), p_time.strftime('%I:%M:%S %p')))
    except Exception:
        conn.rollback()
    finally:
        cursor.close()
        conn.close()
    return recent_punches

def load_attendance_data(today_str):
    conn = psycopg2.connect(DATABASE_URL, sslmode='require')
    cursor = conn.cursor()
    
    query1 = f"""
        SELECT e.emp_code, e.first_name, e.mobile,
               MIN(t.punch_time AT TIME ZONE 'GMT-3') as first_punch,
               MAX(t.punch_time AT TIME ZONE 'GMT-3') as last_punch,
               COUNT(t.id) as punch_count
        FROM personnel_employee e 
        JOIN iclock_transaction t ON e.id = t.emp_id
        WHERE (t.punch_time AT TIME ZONE 'GMT-3')::date = '{today_str}' 
          AND e.emp_code NOT IN ({mgmt_codes_str})
        GROUP BY e.emp_code, e.first_name, e.mobile;
    """
    cursor.execute(query1)
    attendance_rows = cursor.fetchall()
    
    no_out_staff, late_staff = [], []
    for row in attendance_rows:
        emp_code, name, mobile, first_punch, last_punch, punch_count = row
        clean_name = clean_txt(name)
        time_in_clean = first_punch.strftime('%I:%M %p')
        phone_clean = clean_phone(mobile)
        
        if first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15):
            late_staff.append((emp_code, clean_name, phone_clean, time_in_clean))
        
        if punch_count == 1 and not (first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15)):
            no_out_staff.append((emp_code, clean_name, phone_clean, time_in_clean))
            
    query0 = f"""
        SELECT DISTINCT e.emp_code, e.first_name, COALESCE(e.mobile, '') FROM personnel_employee e
        WHERE e.id NOT IN (SELECT DISTINCT emp_id FROM iclock_transaction WHERE (punch_time AT TIME ZONE 'GMT-3')::date = '{today_str}')
          AND e.emp_code NOT IN ({mgmt_codes_str}) 
        ORDER BY e.emp_code ASC;
    """
    cursor.execute(query0)
    full_absent_rows = cursor.fetchall()
    
    full_absent_staff = []
    for row in full_absent_rows:
        if row:
            emp_code, name, mobile = row
            full_absent_staff.append((clean_txt(emp_code), clean_txt(name), clean_phone(mobile)))
    
    cursor.close()
    conn.close()
    return no_out_staff, late_staff, full_absent_staff


# ==========================================
# 3. DASHBOARD INTERFACE LAYOUT RENDERER
# ==========================================
now_syria = datetime.now(SYRIA_TZ)
today_syria_str = now_syria.strftime('%Y-%m-%d')
time_syria_str = now_syria.strftime('%I:%M %p')

# Main Banner Layout
st.title(TEXT_CONFIG["title_main"])
st.header(TEXT_CONFIG["title_sub"])
st.write(TEXT_CONFIG["lbl_date"].format(today_syria_str, time_syria_str))

if st.button(TEXT_CONFIG["btn_refresh"]):
    st.cache_data.clear()

try:
    # --- Live Ticker Logs ---
    st.write("---")
    st.markdown(TEXT_CONFIG["header_live_log"])
    live_logs = load_live_punch_notifications(today_syria_str)
    if live_logs:
        for code, name, p_time in live_logs:
            st.info(TEXT_CONFIG["live_log_row"].format(name, code, p_time))
    else:
        st.caption(TEXT_CONFIG["caption_no_logs"])

    # --- Biometric Devices Layout ---
    st.write("---")
    st.markdown(TEXT_CONFIG["header_devices"])
    devices = load_device_statuses()
    if devices:
        cols = st.columns(len(devices))
        for idx, (alias, status, sn) in enumerate(devices):
            with cols[idx]:
                st.metric(label=f"جهاز: {alias}", value=status, delta=f"SN: {sn[:6]}...")
    else:
        st.warning(TEXT_CONFIG["warn_no_devices"])

    no_out, late, absent = load_attendance_data(today_syria_str)
    st.write("---")
    
    # 1. Late Staff Loop
    st.subheader(TEXT_CONFIG["header_late"].format(len(late)))
    if late:
        for code, name, phone, t_time in late:
            st.write(TEXT_CONFIG["late_row"].format(name, code, t_time))
    else:
        st.success(TEXT_CONFIG["success_no_late"])
        
    st.write("---")
        
    # 2. Absent / Forgot to Punch Loop
    st.subheader(TEXT_CONFIG["header_absent"].format(len(absent)))
    if absent:
        for code, name, phone in absent:
            item_col, action_col = st.columns([4, 1])
            with item_col:
                st.write(TEXT_CONFIG["absent_row"].format(name, code))
            with action_col:
                if phone and phone != '963' and phone != '':
                    sms_text = TEXT_CONFIG["sms_morning"].format(name)
                    encoded_msg = urllib.parse.quote(sms_text)
                    app_url = f"whatsapp://send?phone={phone}&text={encoded_msg}"
                    st.link_button(TEXT_CONFIG["btn_sms_morning"], url=app_url, use_container_width=True)
                else:
                    st.caption(TEXT_CONFIG["caption_no_phone"])
    else:
        st.success(TEXT_CONFIG["success_no_absent"])

    st.write("---")

    # 3. Present Staff / Missing Checkout Loop
    st.subheader(TEXT_CONFIG["header_present"].format(len(no_out)))
    if no_out:
        for code, name, phone, t_time in no_out:
            item_col, action_col = st.columns([4, 1])
            with item_col:
                st.write(TEXT_CONFIG["present_row"].format(name, code, t_time))
            with action_col:
                if now_syria.hour > 18 or (now_syria.hour == 18 and now_syria.minute >= 45):
                    if phone and phone != '963' and phone != '':
                        sms_text = TEXT_CONFIG["sms_evening"].format(name)
                        encoded_msg = urllib.parse.quote(sms_text)
                        app_url = f"whatsapp://send?phone={phone}&text={encoded_msg}"
                        st.link_button(TEXT_CONFIG["btn_sms_evening"], url=app_url, use_container_width=True)
                    else:
                        st.caption(TEXT_CONFIG["caption_no_phone"])
                else:
                    st.caption(TEXT_CONFIG["caption_locked_evening"])
    else:
        st.info(TEXT_CONFIG["info_no_present"])

except Exception as err:
    st.error(TEXT_CONFIG["err_db"].format(err))


