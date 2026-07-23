import streamlit as st
import psycopg2
import pandas as pd
import unicodedata
from datetime import datetime
import zoneinfo
import urllib.parse

# ==========================================
# 1. INITIAL SYSTEM & WINDOW CONFIGURATION
# ==========================================
st.set_page_config(page_title="حضور القصر الذهبي", page_icon="📊", layout="wide")

# Inject clean, universal right-to-left layout alignments for text lines
st.markdown("""
    <style>
    .reportview-container .main .block-container { direction: RTL; text-align: right; }
    h1, h2, h3, h4, p, span, li, div { text-align: right !important; direction: RTL !important; line-height: 1.6 !important; }
    </style>
""", unsafe_allow_html=True)

# System Constants
EXCLUDED_MANAGEMENT_CODES = ("40", "10")
mgmt_codes_str = ",".join(f"'{code}'" for code in EXCLUDED_MANAGEMENT_CODES)
DATABASE_URL = st.secrets["NEON_DATABASE_URL"]

# Explicitly lock the system clock to Syrian time boundaries
SYRIA_TZ = zoneinfo.ZoneInfo("Asia/Damascus")


# ==========================================
# 2. HELPER FUNCTIONS & LIVE DATA SERVICES
# ==========================================
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

def load_attendance_data(today_str):
    conn = psycopg2.connect(DATABASE_URL, sslmode='require')
    cursor = conn.cursor()
    
    # Query 1-Punch and Late Staff
    query1 = f"""
        SELECT e.emp_code, e.first_name, 
               MIN(t.punch_time AT TIME ZONE 'GMT-3') as first_punch,
               MAX(t.punch_time AT TIME ZONE 'GMT-3') as last_punch,
               COUNT(t.id) as punch_count
        FROM personnel_employee e 
        JOIN iclock_transaction t ON e.id = t.emp_id
        WHERE (t.punch_time AT TIME ZONE 'GMT-3')::date = '{today_str}' 
          AND e.emp_code NOT IN ({mgmt_codes_str})
        GROUP BY e.emp_code, e.first_name;
    """
    cursor.execute(query1)
    attendance_rows = cursor.fetchall()
    
    no_out_staff, late_staff = [], []
    for row in attendance_rows:
        emp_code, name, first_punch, last_punch, punch_count = row
        clean_name = clean_txt(name)
        time_in_clean = first_punch.strftime('%I:%M %p')
        
        if first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15):
            late_staff.append((emp_code, clean_name, time_in_clean))
        
        if punch_count == 1 and not (first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15)):
            no_out_staff.append((emp_code, clean_name, time_in_clean))
            
    # Query 0-Punch Staff (Absentees / Forgot to punch)
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
            full_absent_staff.append((clean_txt(emp_code), clean_txt(name), str(mobile).strip()))
    
    cursor.close()
    conn.close()
    return no_out_staff, late_staff, full_absent_staff


# ==========================================
# 3. DASHBOARD INTERFACE LAYOUT RENDERER
# ==========================================
now_syria = datetime.now(SYRIA_TZ)
today_syria_str = now_syria.strftime('%Y-%m-%d')
time_syria_str = now_syria.strftime('%I:%M %p')

# --- 📱 CLEAN NATIVE HEADER BANNER ---
st.title("✨ شركة القصر الذهبي ✨")
st.header("لوحة تحكم إدارة الحضور والغياب")
st.write(f"📅 التاريخ: **{today_syria_str}**  │  ⏰ الوقت الحالي في سوريا: **{time_syria_str}**")

if st.button("🔄 تحديث البيانات الحية الآن"):
    st.cache_data.clear()

try:
    # --- 📠 LIVE HARDWARE COUNTER DASHBOARD ---
    st.write("---")
    st.markdown("### 📡 حالة اتصال أجهزة البصمة الحالية:")
    devices = load_device_statuses()
    
    if devices:
        cols = st.columns(len(devices))
        for idx, (alias, status, sn) in enumerate(devices):
            with cols[idx]:
                st.metric(label=f"جهاز: {alias}", value=status, delta=f"SN: {sn[:6]}...")
    else:
        st.warning("⚠️ لا توجد أجهزة مضافة أو تعذر تحميل البيانات.")

    no_out, late, absent = load_attendance_data(today_syria_str)
    st.write("---")
    
    # 1. Render Late Staff Section
    st.subheader(f"⏰ المتأخرون اليوم ({len(late)}) – دخول بعد 09:15 صباحاً")
    if late:
        for code, name, t_time in late:
            st.write(f"🔸 **{name}** (كود: {code}) ── وقت الدخول: {t_time}")
    else:
        st.success("🎉 لا يوجد متأخرين اليوم!")
        
    st.write("---")
        
    # 2. UPDATED: Removed English text from header
    st.subheader(f"❌ غائبون أو نسوا تسجيل الحضور ({len(absent)})")
    if absent:
        for code, name, mobile in absent:
            item_col, action_col = st.columns([4, 1])
            with item_col:
                st.write(f"🔹 **{name}** (كود: {code})")
            with action_col:
                if mobile:
                    phone_formatted = mobile if mobile.startswith('+') or len(mobile) > 10 else f"963{mobile.lstrip('0')}"
                    msg = f"مرحباً {name}، يرجى العلم أنه لم يتم تسجيل بصمة حضور لك اليوم المندرج بتاريخ {today_syria_str}. إذا كنت متواجداً بالعمل، يرجى مراجعة الإدارة أو تأكيد البصمة مسبقاً."
                    encoded_msg = urllib.parse.quote(msg)
                    wa_url = f"https://wa.me{phone_formatted}?text={encoded_msg}"
                    st.markdown(f'<a href="{wa_url}" target="_blank" style="text-decoration:none;"><button style="background-color:#25D366; color:white; border:none; padding:4px 10px; border-radius:4px; cursor:pointer;">💬 مراسلة تذكيرية</button></a>', unsafe_allow_html=True)
                else:
                    st.caption("🚫 لا يوجد رقم")
    else:
        st.success("🎉 لا يوجد غيابات اليوم!")

    st.write("---")

    # 3. UPDATED: Removed English text from header
    st.subheader(f"🟢 الموظفون المتواجدون حالياً في العمل ({len(no_out)})")
    if no_out:
        for code, name, t_time in no_out:
            st.write(f"🔸 **{name}** (كود: {code}) ── وقت الدخول: {t_time}")
    else:
        st.info("لا يوجد موظفين منتظمين متواجدين حالياً.")

except Exception as err:
    st.error(f"خطأ في الاتصال بقاعدة البيانات السحابية: {err}")
