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
st.set_page_config(page_title="Golden Palace Attendance", page_icon="📊", layout="wide")

# Inject clean, universal left-to-right layout alignments for English text
st.markdown("""
    <style>
    .reportview-container .main .block-container { direction: LTR; text-align: left; }
    h1, h2, h3, h4, p, span, li, div { text-align: left !important; direction: LTR !important; line-height: 1.6 !important; }
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
            status_tag = "🟢 Online" if is_online and (str(is_online).strip().lower() in ('true', '1', 't', 'y', 'yes')) else "🔴 Offline"
            device_metrics.append((clean_txt(alias), status_tag, sn))
    except Exception:
        conn.rollback()
        try:
            query = "SELECT alias, last_activity, sn FROM iclock_terminal;"
            cursor.execute(query)
            rows = cursor.fetchall()
            
            timestamps = [r[1] for r in rows if r and r[1]]
            latest_system_ping = max(timestamps) if timestamps else None
            
            for row in rows:
                alias, last_act, sn = row
                if last_act and latest_system_ping:
                    seconds_elapsed = (latest_system_ping.replace(tzinfo=None) - last_act.replace(tzinfo=None)).total_seconds()
                    status_tag = "🟢 Online" if seconds_elapsed < 600 else "🔴 Offline"
                else:
                    status_tag = "🔴 Offline"
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
        
        # Check Late Arrival (After 09:15 AM)
        if first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15):
            late_staff.append((emp_code, clean_name, time_in_clean))
        
        # Check Current Presence: Odd number of punches means they are clocked in and still on-site
        if punch_count % 2 != 0 or first_punch == last_punch:
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
st.title("✨ Golden Palace Co. ✨")
st.header("Attendance Management Dashboard")
st.write(f"📅 Date: **{today_syria_str}**  │  ⏰ Current Time (Syria): **{time_syria_str}**")

if st.button("🔄 Refresh Live Data Now"):
    st.cache_data.clear()

try:
    # --- 📠 LIVE HARDWARE COUNTER DASHBOARD ---
    st.write("---")
    st.markdown("### 📡 Biometric Device Status:")
    devices = load_device_statuses()
    
    if devices:
        cols = st.columns(len(devices))
        for idx, (alias, status, sn) in enumerate(devices):
            with cols[idx]:
                st.metric(label=f"Device: {alias}", value=status, delta=f"SN: {sn[:6]}...")
    else:
        st.warning("⚠️ No devices detected or connection could not be established.")

    no_out, late, absent = load_attendance_data(today_syria_str)
    st.write("---")
    
    # 1. Render Late Staff Section
    st.subheader(f"⏰ Late Arrivals Today ({len(late)}) – Checked in after 09:15 AM")
    if late:
        for code, name, t_time in late:
            st.write(f"🔸 **{name}** (Code: {code}) ── Check-in Time: {t_time}")
    else:
        st.success("🎉 No late arrivals today!")
        
    st.write("---")
        
    # 2. Render Absent / Forgot to punch section
    st.subheader(f"❌ Absent / Missing Punch Records ({len(absent)})")
    if absent:
        for code, name, mobile in absent:
            item_col, action_col = st.columns([5, 1])
            with item_col:
                st.write(f"🔹 **{name}** (Code: {code})")
            with action_col:
                if mobile and mobile != 'None' and mobile != '':
                    clean_phone = mobile.lstrip('+').lstrip('0')
                    phone_formatted = clean_phone if clean_phone.startswith('963') else f"963{clean_phone}"
                    
                    msg = f"Hello {name}, please note that no biometric check-in log was recorded for you today ({today_syria_str}). If you are currently at work, please visit HR or confirm your punch records."
                    encoded_msg = urllib.parse.quote(msg)
                    wa_url = f"https://wa.me{phone_formatted}?text={encoded_msg}"
                    
                    st.link_button("💬 Remind", url=wa_url, use_container_width=True)
                else:
                    st.caption("🚫 No Phone Number")
    else:
        st.success("🎉 No absences tracked today!")

    st.write("---")

    # 3. Render Present Staff Section
    st.subheader(f"🟢 Staff Currently Active On-Site ({len(no_out)})")
    if no_out:
        for code, name, t_time in no_out:
            st.write(f"🔸 **{name}** (Code: {code}) ── Check-in Time: {t_time}")
    else:
        st.info("No active staff records currently marked on site.")

except Exception as err:
    st.error(f"Cloud database connection error: {err}")
