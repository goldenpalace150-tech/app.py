import streamlit as st
import psycopg2
import pandas as pd
import unicodedata
from datetime import datetime
import zoneinfo

# Configure the mobile webpage title and centered wide layout
st.set_page_config(page_title="حضور القصر الذهبي", page_icon="📊", layout="wide")

# Inject custom right-to-left CSS styling to guarantee beautiful Arabic text rendering
st.markdown("""
    <style>
    .reportview-container .main .block-container { direction: RTL; unicode-bidi: bidi-override; text-align: right; }
    h1, h2, h3, p, span, li, div { text-align: right !important; direction: RTL !important; }
    .stDataFrame { direction: RTL !important; text-align: right !important; }
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

def load_attendance_data(today_str):
    conn = psycopg2.connect(DATABASE_URL, sslmode='require')
    cursor = conn.cursor()
    
    # 1. Query 1-Punch and Late Staff (Locked to Syrian time conversion)
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
    full_absent_staff = [(r, clean_txt(r)) for r in full_absent_rows]
    
    cursor.close()
    conn.close()
    return no_out_staff, late_staff, full_absent_staff

# Forces the current header time to evaluate directly using the Damascus clock
now_syria = datetime.now(SYRIA_TZ)
today_syria_str = now_syria.strftime('%Y-%m-%d')
time_syria_str = now_syria.strftime('%I:%M %p')

# --- 📱 HEADER WITH GOLDEN PALACE BRAND LOGO ---
# FIXED: Replaced Syrian flag with your high-resolution hosted company logo vector
st.markdown(
    f"""
    <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; margin-bottom: 25px; gap: 10px;">
        <img src="https://ibb.co" width="220" style="margin-bottom: 5px;">
        <h2 style="margin: 0; padding: 0; color: #D4AF37; font-weight: bold;">لوحة تحكم إدارة الحضور والغياب</h2>
        <h4 style="margin: 0; padding: 0; color: #555;">تاريخ اليوم: {today_syria_str} | التوقيت الحالي في سوريا: {time_syria_str}</h4>
    </div>
    """, 
    unsafe_allow_html=True
)

if st.button("🔄 تحديث البيانات الحية الآن"):
    st.cache_data.clear()

try:
    no_out, late, absent = load_attendance_data(today_syria_str)
    st.write("---")
    
    # Render Late Staff Section
    st.markdown(f"### ⏰ المتأخرون اليوم ({len(late)}) – بصمة دخول بعد 09:15 صباحاً")
    if late:
        for code, name, t_time in late:
            st.markdown(f"🔸 **{name}** (كود: {code}) ── 🕛 وقت الدخول: **{t_time}**")
    else:
        st.success("🎉 لا يوجد متأخرين اليوم!")
        
    st.write("---")
        
    # Render Absent Section
    st.markdown(f"### ❌ غائبون تماماً اليوم ({len(absent)}) – 0 بصمة")
    if absent:
        for code, name in absent:
            st.markdown(f"🔹 **{name}** (كود: {code})")
    else:
        st.success("🎉 لا يوجد غيابات كاملة اليوم!")

    st.write("---")

    # Render Normal 1-Punch Section
    st.markdown(f"### ⚠️ سجلوا دخول في الوقت ولم يسجلوا خروج بعد ({len(no_out)})")
    if no_out:
        for code, name, t_time in no_out:
            st.markdown(f"🔸 **{name}** (كود: {code}) ── 🕒 وقت الدخول: **{t_time}**")
    else:
        st.markdown("لا يوجد موظفين منتظمين بانتظار الخروج.")

except Exception as err:
    st.error(f"خطأ في الاتصال بقاعدة البيانات السحابية: {err}")
