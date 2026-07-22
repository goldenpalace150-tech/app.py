import streamlit as st
import psycopg2
import pandas as pd
import unicodedata
from datetime import datetime, timedelta
import zoneinfo

# Configure the mobile webpage title and centered layout
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

def clean_txt(raw_text):
    if not raw_text: return ""
    return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip())

def load_device_statuses():
    """Queries Neon to check exactly which machines are online based on their last ping"""
    conn = psycopg2.connect(DATABASE_URL, sslmode='require')
    cursor = conn.cursor()
    
    # Query your specific terminal model metadata parameters
    query = "SELECT alias, last_activity, sn FROM iclock_terminal;"
    cursor.execute(query)
    rows = cursor.fetchall()
    
    device_metrics = []
    now_ Damascus = datetime.now(SYRIA_TZ)
    
    for row in rows:
        alias, last_act, sn = row
        clean_alias = clean_txt(alias)
        
        # If the device has communicated within the last 5 minutes, it is strictly Online
        if last_act and (now_Damascus.replace(tzinfo=None) - last_act).total_seconds() < 300:
            status_tag = "🟢 متصل"
        else:
            status_tag = "🔴 غير متصل"
            
        device_metrics.append((clean_alias, status_tag, sn))
        
    cursor.close()
    conn.close()
    return device_metrics

def load_attendance_data(today_str):
    conn = psycopg2.connect(DATABASE_URL, sslmode='require')
    cursor = conn.cursor()
    
    # 1. Query 1-Punch and Late Staff 
    query1 = f"""
        SELECT DISTINCT e.emp_code, e.first_name, MIN(t.punch_time AT TIME ZONE 'GMT-3') 
        FROM personnel_employee e JOIN iclock_transaction t ON e.id = t.emp_id
        WHERE (t.punch_time AT TIME ZONE 'GMT-3')::date = '{today_str}' AND e.emp_code NOT IN ({mgmt_codes_str})
        GROUP BY e.emp_code, e.first_name;
    """
    cursor.execute(query1)
    one_punch_rows = cursor.fetchall()
    
    no_out_staff, late_staff = [], []
    for row in one_punch_rows:
        emp_code, name, p_time = row
        clean_name = clean_txt(name)
        time_clean = p_time.strftime('%I:%M %p')
        
        if p_time.hour > 9 or (p_time.hour == 9 and p_time.minute > 15):
            late_staff.append((emp_code, clean_name, time_clean))
        else:
            no_out_staff.append((emp_code, clean_name, time_clean))
            
    # 2. Query 0-Punch Staff (Absentees)
    query0 = f"""
        SELECT DISTINCT e.emp_code, e.first_name FROM personnel_employee e
        WHERE e.id NOT IN (SELECT DISTINCT emp_id FROM iclock_transaction WHERE (punch_time AT TIME ZONE 'GMT-3')::date = '{today_str}')
        AND e.emp_code NOT IN ({mgmt_codes_str}) ORDER BY e.emp_code ASC;
    """
    cursor.execute(query0)
    full_absent_rows = cursor.fetchall()
    full_absent_staff = [(row, clean_txt(row)) for row in full_absent_rows if row]
    
    cursor.close()
    conn.close()
    return no_out_staff, late_staff, full_absent_staff

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
    
    # Render hardware states dynamically in a clean row format
    cols = st.columns(len(devices))
    for idx, (alias, status, sn) in enumerate(devices):
        with cols[idx]:
            st.metric(label=f"جهاز: {alias}", value=status, delta=f"SN: {sn[:6]}...")

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
        
    # 2. Render Absent Section
    st.subheader(f"❌ غائبون تماماً اليوم ({len(absent)}) – 0 بصمة")
    if absent:
        for code, name in absent:
            st.write(f"🔹 **{name}** (كود: {code})")
    else:
        st.success("🎉 لا يوجد غيابات كاملة اليوم!")

    st.write("---")

    # 3. Render Normal 1-Punch Section
    st.subheader(f"⚠️ سجلوا دخول ولم يسجلوا خروج بعد ({len(no_out)})")
    if no_out:
        for code, name, t_time in no_out:
            st.write(f"🔸 **{name}** (كود: {code}) ── وقت الدخول: {t_time}")
    else:
        st.info("لا يوجد موظفين منتظمين بانتظار الخروج.")

except Exception as err:
    st.error(f"خطأ في الاتصال بقاعدة البيانات السحابية: {err}")
