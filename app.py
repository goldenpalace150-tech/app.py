import sys
import time
import threading
import psycopg2
import subprocess
import pyautogui
import urllib.parse
from datetime import datetime
import streamlit as st

# ==========================================
# 1. LOCAL OFFICE DATABASE ONLY CONFIGURATION
# ==========================================
LOCAL_DB = {
    "user": "postgres",
    "password": "nevermind.123",  
    "host": "127.0.0.1",
    "database": "biotime",
    "port": 7496
}

# Management tracking exclusions for the visual lists
EXCLUDED_CODES = ("40", "10", "20")

# ==========================================
# 2. LOCAL AUTOMATION CONTROLLER (WHATSAPP)
# ==========================================
def execute_whatsapp_keystroke(phone, message):
    """Automates typing commands directly inside your logged-in desktop application."""
    try:
        time.sleep(5)
        process = subprocess.Popen(['clip'], stdin=subprocess.PIPE, close_fds=True)
        process.communicate(input=message.encode('utf-16'))
        
        cmd = f'start "" "whatsapp://send?phone={phone}"'
        subprocess.run(cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(1.5)
        
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.3)
        pyautogui.press('enter')
        print(f"🚀 Locally automated message sent to: {phone}")
    except Exception as e:
        print(f"❌ Automation runtime error: {e}")

# ==========================================
# 3. BACKGROUND PROCESSING ENGINE
# ==========================================
def background_punch_monitor():
    """Monitors local database changes and processes live messaging flags automatically."""
    print("📡 Local background live monitoring engine active...")
    try:
        conn = psycopg2.connect(**LOCAL_DB)
        cursor = conn.cursor()
    except Exception as e:
        print(f"❌ Database initialization failed: {e}")
        return

    # Track from the current max transaction ID onward to prevent historic spamming
    cursor.execute("SELECT MAX(id) FROM iclock_transaction;")
    res = cursor.fetchone()
    last_processed_id = res[0] if res and res[0] is not None else 0
    
    while True:
        try:
            today_str = datetime.now().strftime('%Y-%m-%d')
            
            # Fetch new punches directly from local memory tables adjusted for Syria
            query = """
                SELECT t.id, e.emp_code, e.first_name, e.mobile, (t.punch_time AT TIME ZONE 'GMT-3')
                FROM iclock_transaction t
                JOIN personnel_employee e ON t.emp_id = e.id
                WHERE t.id > %s
                ORDER BY t.id ASC;
            """
            cursor.execute(query, (last_processed_id,))
            new_punches = cursor.fetchall()
            
            for punch in new_punches:
                t_id, emp_code, first_name, mobile, punch_time = punch
                time_str = punch_time.strftime('%I:%M:%S %p')
                
                if not mobile or str(mobile).strip() == "":
                    last_processed_id = t_id
                    continue
                
                # Check punch counts to determine state
                count_query = """
                    SELECT COUNT(id) FROM iclock_transaction 
                    WHERE emp_id = (SELECT id FROM personnel_employee WHERE emp_code = %s)
                    AND (punch_time AT TIME ZONE 'GMT-3')::date = %s AND id <= %s;
                """
                cursor.execute(count_query, (emp_code, today_str, t_id))
                res_count = cursor.fetchone()
                punch_count = res_count[0] if res_count else 1
                
                if punch_count % 2 != 0:
                    msg = f"مرحباً {first_name}، تم تسجيل بصمة *الدخول* بنجاح عند الساعة {time_str}. أتمنى لك يوماً سعيداً! ✨"
                else:
                    msg = f"مرحباً {first_name}، تم تسجيل بصمة *الخروج* بنجاح عند الساعة {time_str}. رافقتك السلامة! 🏡"
                
                # Execute direct localized messaging routine
                execute_whatsapp_keystroke(mobile, msg)
                last_processed_id = t_id
                
            time.sleep(5)
        except Exception:
            time.sleep(5)

# ==========================================
# 4. STREAMLIT VISUAL DASHBOARD INTERFACE
# ==========================================
def run_dashboard_ui():
    st.set_page_config(page_title="حضور القصر الذهبي", page_icon="📊", layout="wide")
    
    # Force RTL Arabic Formatting Styles
    st.markdown("""
        <style>
        .reportview-container .main .block-container { direction: RTL; text-align: right; }
        h1, h2, h3, h4, p, span, li, div { text-align: right !important; direction: RTL !important; line-height: 1.6 !important; }
        </style>
    """, unsafe_allow_html=True)
    
    st.title("✨ شركة القصر الذهبي ✨")
    st.subheader("لوحة تحكم إدارة الحضور والغياب (نظام محلي موحد)")
    
    now = datetime.now()
    st.markdown(f"📅 التاريخ: **{now.strftime('%Y-%m-%d')}**  │  ⏰ الوقت الحالي في سوريا: **{now.strftime('%I:%M:%S %p')}**")
    
    if st.button("🔄 تحديث البيانات الآن"):
        st.rerun()

    # Load layout data directly from your local Postgres socket instances
    try:
        conn = psycopg2.connect(**LOCAL_DB)
        cursor = conn.cursor()
        today_str = now.strftime('%Y-%m-%d')
        mgmt_str = ",".join(f"'{c}'" for c in EXCLUDED_CODES)

        # 1. Query Present and Late metrics with the GMT-3 conversion patch
        cursor.execute(f"""
            SELECT e.emp_code, e.first_name, e.mobile, MIN(t.punch_time AT TIME ZONE 'GMT-3') as first_punch, COUNT(t.id) as p_count
            FROM personnel_employee e 
            JOIN iclock_transaction t ON e.id = t.emp_id
            WHERE (t.punch_time AT TIME ZONE 'GMT-3')::date = '{today_str}' AND e.emp_code NOT IN ({mgmt_str})
            GROUP BY e.emp_code, e.first_name, e.mobile;
        """)
        attendance_rows = cursor.fetchall()
        
        late_staff, present_staff = [], []
        for row in attendance_rows:
            emp_code, name, mobile, first_punch, p_count = row
            time_in = first_punch.strftime('%I:%M %p')
            
            if first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15):
                late_staff.append((emp_code, name, mobile, time_in))
            if p_count % 2 != 0:
                present_staff.append((emp_code, name, mobile, time_in))

        # 2. Query Full Absent Metrics
        cursor.execute(f"""
            SELECT DISTINCT e.emp_code, e.first_name FROM personnel_employee e
            WHERE e.id NOT IN (
                SELECT DISTINCT emp_id FROM iclock_transaction WHERE (punch_time AT TIME ZONE 'GMT-3')::date = '{today_str}'
            ) AND e.emp_code NOT IN ({mgmt_str}) ORDER BY e.emp_code ASC;
        """)
        absent_rows = cursor.fetchall()

        # Render Interface Grid Columns
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown(f"### ⏰ المتأخرون اليوم ({len(late_staff)})")
            if late_staff:
                for emp_code, name, mobile, time_in in late_staff:
                    st.write(f"🔸 **{name}** (كود: {emp_code}) ── الدخول: {time_in}")
            else:
                st.success("🎉 لا يوجد متأخرين اليوم!")

        with col2:
            st.markdown(f"### 🟢 المتواجدون حالياً ({len(present_staff)})")
            if present_staff:
                for emp_code, name, mobile, time_in in present_staff:
                    st.write(f"🔸 **{name}** (كود: {emp_code}) ── الدخول: {time_in}")
            else:
                st.info("لا يوجد موظفين متواجدين حالياً.")

        with col3:
            st.markdown(f"### ❌ غائبون أو نسوا البصمة ({len(absent_rows)})")
            if absent_rows:
                for emp_code, name in absent_rows:
                    st.write(f"🔹 **{name}** (كود: {emp_code})")
            else:
                st.success("🎉 لا يوجد غيابات اليوم!")
                
        cursor.close()
        conn.close()
    except Exception as err:
        st.error(f"خطأ في الاتصال بقاعدة البيانات المحلية: {err}")

# ==========================================
# 5. MONOLITHIC EXECUTION ENTRY POINT
# ==========================================
if __name__ == "__main__":
    if "monitor_started" not in st.session_state:
        threading.Thread(target=background_punch_monitor, daemon=True).start()
        st.session_state["monitor_started"] = True
        
    run_dashboard_ui()
