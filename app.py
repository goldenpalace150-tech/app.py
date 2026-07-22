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

# NATIVE BASE64 DATA ENCODING FOR GOLDEN PALACE LOGO
# This keeps the logo fully intact inside the file without requiring web link loaders
LOGO_BASE64 = (
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAcIAAABwCAMAAACZ70idAAAAAXNSR0IArs4c6QAAA"
    "FAGQ0hREQYMSDExITExMTExMTExMTExMTExMTExMTExMTExMTExMTExMTExMTExMTExMTExMTExMTExMTExM"
    "TExMTExMTExMTExMTExMTEhId0F28gAAAAd0SU1FBmUHFgwwDBYwFiwAAALwSURBVGje7ZjbYpswEEVpGgK"
    "BQP//l08OtoMscYgTY6+896m1Z63pWwYymCHYvYgAAAAA8M/g6eP+q+v7G2m7vY+A668S3UjYvY+A66wS6Ub"
    "C7n0E8v4G791I2P2fRL8Sdu8jkPcbfpWwex+B3e5v8atENxL2fRLpRsLufQT2X8H0K2H3PgLpRsLufQTM9y"
    "M63UjYvY+A668S3UjYvY9A3m/47f7bS9i9j0C7v8WvEt1I2PdJwtxI2L2PANv3S/Rlwu59BPD+S/S6kbD7k"
    "8DpfwS3m9FvL2H3JwFrP0XvO9JuRs+NhH2fhK//EH9N7K8S3UzY/UnA9b8G68yG3UzY/UngR07fJfo6YV8n"
    "COfG8z6N0U+Npx8Z/fREHxn98Yl+eCKNThDei6fpxZMfE+mXie7F858Zff/Y6PtHRuofGen7R/fijU7wYvI"
    "3nvwKxr+C8StYv4LpT8b0R2N8bIz3jW7F++IEL6bU6MvE+mVi+hMxXSbWPxLrr2DKf7NidIIPU2r6b8X4f7"
    "NidIIPU2r6b8V6mdjPivEysX6ZmN6K8WfFepnYK8aeE/vF7ASfptTsPzX+/p/XOf6p0e8b/fREpBM8TfnUe"
    "Pr+8X6ZmO5PxF4T++REpBN8mhP9f2a8Xya6E/sZMd6fifEnI53gaU6N94v978R0JxO9T07wPMH0byL2byL2"
    "itGdjN7fTfB9Svy9f7xfJp7mBO9To+9Tou8f779E7BOj7x+f93wSPk3pE3vFpzlB+G/CpxOETyd4OsHTCZ5"
    "O8HSCpxM8neDpBE8neDoBAAAAAAAAWAn/ABq3O9V7N86YAAAAAElFTkSuQmCC"
)

def clean_txt(raw_text):
    if not raw_text: return ""
    return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip())

def load_attendance_data(today_str):
    conn = psycopg2.connect(DATABASE_URL, sslmode='require')
    cursor = conn.cursor()
    
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

now_syria = datetime.now(SYRIA_TZ)
today_syria_str = now_syria.strftime('%Y-%m-%d')
time_syria_str = now_syria.strftime('%I:%M %p')

# --- 📱 HEADER WITH EMBEDDED INTERNAL LOGO ---
st.markdown(
    f"""
    <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; margin-bottom: 25px; gap: 10px;">
        <img src="{LOGO_BASE64}" width="240" style="margin-bottom: 5px;">
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
    
    st.markdown(f"### ⏰ المتأخرون اليوم ({len(late)}) – بصمة دخول بعد 09:15 صباحاً")
    if late:
        for code, name, t_time in late:
            st.markdown(f"🔸 **{name}** (كود: {code}) ── 🕛 وقت الدخول: **{t_time}**")
    else:
        st.success("🎉 لا يوجد متأخرين اليوم!")
        
    st.write("---")
        
    st.markdown(f"### ❌ غائبون تماماً اليوم ({len(absent)}) – 0 بصمة")
    if absent:
        for code, name in absent:
            st.markdown(f"🔹 **{name}** (كود: {code})")
    else:
        st.success("🎉 لا يوجد غيابات كاملة اليوم!")

    st.write("---")

    st.markdown(f"### ⚠️ سجلوا دخول في الوقت ولم يسجلوا خروج بعد ({len(no_out)})")
    if no_out:
        for code, name, t_time in no_out:
            st.markdown(f"🔸 **{name}** (كود: {code}) ── 🕒 وقت الدخول: **{t_time}**")
    else:
        st.markdown("لا يوجد موظفين منتظمين بانتظار الخروج.")

except Exception as err:
    st.error(f"خطأ في الاتصال بقاعدة البيانات السحابية: {err}")
