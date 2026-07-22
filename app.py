import streamlit as st
import psycopg2
import pandas as pd
import unicodedata
from datetime import datetime

# Configure the mobile webpage title and clean presentation layout
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
EXCLUDED_MANAGEMENT_CODES = ("40", "10", "20")
mgmt_codes_str = ",".join(f"'{code}'" for code in EXCLUDED_MANAGEMENT_CODES)

def clean_txt(raw_text):
    if not raw_text: return ""
    return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip())

# Fetch variables securely from Streamlit's built-in cloud environment vault
DATABASE_URL = st.secrets["NEON_DATABASE_URL"]

def load_attendance_data():
    today_str = datetime.now().strftime('%Y-%m-%d')
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
        # Format the exact punch time cleanly
        time_clean = p_time.strftime('%I:%M %p')
        
        # Merge time directly into the text display label string
        display_label = f"{clean_name} ({time_clean})"
        
        if p_time.hour > 9 or (p_time.hour == 9 and p_time.minute > 15):
            late_staff.append({"الكود": emp_code, "الموظف ووقت الدخول": display_label})
        else:
            no_out_staff.append({"الكود": emp_code, "الموظف ووقت الدخول": display_label})

        else:
            no_out_staff.append({"الكود": emp_code, "الاسم": clean_name, "وقت الدخول": time_clean})
            
    # 2. Query 0-Punch Staff (Absentees)
    query0 = f"""
        SELECT DISTINCT e.emp_code, e.first_name FROM personnel_employee e
        WHERE e.id NOT IN (SELECT DISTINCT emp_id FROM iclock_transaction WHERE (punch_time AT TIME ZONE 'GMT-3')::date = '{today_str}')
        AND e.emp_code NOT IN ({mgmt_codes_str}) ORDER BY e.emp_code ASC;
    """
    cursor.execute(query0)
    full_absent_rows = cursor.fetchall()
    full_absent_staff = [{"الكود": r[0], "الاسم": clean_txt(r[1])} for r in full_absent_rows]
    
    cursor.close()
    conn.close()
    return no_out_staff, late_staff, full_absent_staff

# --- 📱 MOBILE WEB INTERFACE GRAPHICS DISPLAY ---
st.title("🏆 لوحة تحكم شركة القصر الذهبي")
st.subheader(f"تاريخ اليوم: {datetime.now().strftime('%Y-%m-%d')} | التوقيت الحالي: {datetime.now().strftime('%I:%M %p')}")

if st.button("🔄 تحديث البيانات الحية الآن"):
    st.cache_data.clear()

try:
    no_out, late, absent = load_attendance_data()
    
    st.write("---")
    
    # Render Late Staff Section
    st.error(f"⏰ المتأخرون اليوم ({len(late)}) - بصمة دخول بعد 09:15 صباحاً")
    if late: st.dataframe(pd.DataFrame(late), use_container_width=True, hide_index=True)
    else: st.success("🎉 لا يوجد متأخرين اليوم!")
        
    # Render Absent Section
    st.warning(f"❌ غائبون تماماً اليوم ({len(absent)}) - 0 بصمة")
    if absent: st.dataframe(pd.DataFrame(absent), use_container_width=True, hide_index=True)
    else: st.success("🎉 لا يوجد غيابات كاملة اليوم!")

    # Render Normal 1-Punch Section
    st.info(f"⚠️ سجلوا دخول ولم يسجلوا خروج بعد ({len(no_out)})")
    if no_out: st.dataframe(pd.DataFrame(no_out), use_container_width=True, hide_index=True)
    else: st.write("لا يوجد موظفين منتظمين بانتظار الخروج.")

except Exception as err:
    st.error(f"خطأ في الاتصال بقاعدة البيانات السحابية: {err}")
