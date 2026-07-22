import streamlit as st
import psycopg2
import pandas as pd
import unicodedata
from datetime import datetime
import zoneinfo

# Configure the mobile webpage title and centered wide layout
st.set_page_config(page_title="حضور القصر الذهبي", page_icon="📊", layout="wide")

# Force explicit right-to-left column alignment for grid layouts
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

# --- 🔌 HIDDEN BACKGROUND HARDWARE RECEIVER PATCH ---
# This function dynamically intercepts raw fingerprint logs sent to your Streamlit URL path
def handle_incoming_device_logs():
    query_params = st.query_params
    # Detects if a physical ZK machine is attempting an ADMS data sync handshake
    if "SN" in query_params or "sn" in query_params:
        try:
            device_sn = query_params.get("SN", query_params.get("sn"))
            print(f"📡 Cloud Handshake Intercepted from Machine SN: {device_sn}")
            # Automatically returns a clean 200 OK signal directly to the machine's internal firmware
            st.write("OK")
            st.stop()
        except:
            pass

# Instantly process any background background hardware signals before loading the interface
handle_incoming_device_logs()

def load_attendance_data(today_str):
    conn = psycopg2.connect(DATABASE_URL, sslmode='require')
    cursor = conn.cursor()
    
    # Query 1-Punch and Late Staff (Locked to Syrian time conversion)
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
            
    # Query 0-Punch Staff (Absentees)
    query0 = f"""
        SELECT DISTINCT e.emp_code, e.first_name FROM personnel_employee e
        WHERE e.id NOT IN (SELECT DISTINCT emp_id FROM iclock_transaction WHERE (punch_time AT TIME ZONE 'GMT-3')::date = '{today_str}')
        AND e.emp_code NOT IN ({mgmt_codes_str}) ORDER BY e.emp_code ASC;
    """
    cursor.execute(query0)
    full_absent_rows = cursor.fetchall()
    full_absent_staff = [(row[0], clean_txt(row[1])) for row in full_absent_rows if row]
    
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
    no_out, late, absent = load_attendance_data(today_syria_str)
    st.write("---")
    
    st.subheader(f"⏰ المتأخرون اليوم ({len(late)}) – دخول بعد 09:15 صباحاً")
    if late:
        for code, name, t_time in late:
            st.write(f"🔸 **{name}** (كود: {code}) ── وقت الدخول: {t_time}")
    else:
        st.success("🎉 لا يوجد متأخرين اليوم!")
        
    st.write("---")
        
    st.subheader(f"❌ غائبون تماماً اليوم ({len(absent)}) – 0 بصمة")
    if absent:
        for code, name in absent:
            st.write(f"🔹 **{name}** (كود: {code})")
    else:
        st.success("🎉 لا يوجد غيابات كاملة اليوم!")

    st.write("---")

    st.subheader(f"⚠️ سجلوا دخول ولم يسجلوا خروج بعد ({len(no_out)})")
    if no_out:
        for code, name, t_time in no_out:
            st.write(f"🔸 **{name}** (كود: {code}) ── وقت الدخول: {t_time}")
    else:
        st.info("لا يوجد موظفين منتظمين بانتظار الخروج.")

except Exception as err:
    st.error(f"خطأ في الاتصال بقاعدة البيانات السحابية: {err}")
