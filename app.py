import streamlit as st
import psycopg2
import unicodedata
from datetime import datetime
import zoneinfo
import urllib.parse

# ==========================================
# 0. RTL ARABIC TEXT CONSTANTS
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
    "header_late": "⏰ المتأخرون اليوم ({}) – دخول بعد 09:15 صباحاً",
    "late_row": "🔸 **{}** (كود: {}) ── وقت الدخول: {}",
    "success_no_late": "🎉 لا يوجد متأخرين اليوم!",
    "header_absent": "❌ غائبون أو نسوا تسجيل الحضور ({})",
    "absent_row": "🔹 **{}** (كود: {})",
    "success_no_absent": "🎉 لا يوجد غيابات اليوم!",
    "header_present": "🟢 الموظفون المتواجدون حالياً في العمل ({})",
    "present_row": "🔸 **{}** (كود: {}) ── وقت الدخول: {}",
    "info_no_present": "لا يوجد موظفين منتظمين متواجدين حالياً.",
    "err_db": "خطأ في الاتصال بقاعدة البيانات: {}"
}

st.set_page_config(page_title=TEXT_CONFIG["page_title"], page_icon="📊", layout="wide")
st.markdown(TEXT_CONFIG["style_align"], unsafe_allow_html=True)

# Exclude management codes 10, 20, 40 from dashboard display
EXCLUDED_MANAGEMENT_CODES = ("40", "10", "20")
mgmt_codes_str = ",".join(f"'{code}'" for code in EXCLUDED_MANAGEMENT_CODES)
DATABASE_URL = st.secrets["NEON_DATABASE_URL"]
SYRIA_TZ = zoneinfo.ZoneInfo("Asia/Damascus")

def clean_txt(raw_text):
    if not raw_text: return ""
    return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip())

def load_attendance_data(today_str):
    conn = psycopg2.connect(DATABASE_URL)
    cursor = conn.cursor()
    
    query1 = f"""
        SELECT e.emp_code, e.first_name, 
               MIN(t.punch_time AT TIME ZONE 'GMT-3') as first_punch,
               COUNT(t.id) as punch_count
        FROM personnel_employee e 
        JOIN iclock_transaction t ON e.id = t.emp_id
        WHERE (t.punch_time AT TIME ZONE 'GMT-3')::date = '{today_str}' 
          AND e.emp_code NOT IN ({mgmt_codes_str})
        GROUP BY e.emp_code, e.first_name;
    """
    cursor.execute(query1)
    attendance_rows = cursor.fetchall()
    
    present_staff, late_staff = [], []
    for row in attendance_rows:
        emp_code, name, first_punch, punch_count = row
        clean_name = clean_txt(name)
        time_in_clean = first_punch.strftime('%I:%M %p')
        
        if first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15):
            late_staff.append((emp_code, clean_name, time_in_clean))
        
        if punch_count % 2 != 0:
            present_staff.append((emp_code, clean_name, time_in_clean))
            
    query0 = f"""
        SELECT DISTINCT e.emp_code, e.first_name FROM personnel_employee e
        WHERE e.id NOT IN (SELECT DISTINCT emp_id FROM iclock_transaction WHERE (punch_time AT TIME ZONE 'GMT-3')::date = '{today_str}')
          AND e.emp_code NOT IN ({mgmt_codes_str}) 
        ORDER BY e.emp_code ASC;
    """
    cursor.execute(query0)
    full_absent_rows = cursor.fetchall()
    
    full_absent_staff = []
    for row in full_absent_rows:
        if row:
            emp_code, name = row
            full_absent_staff.append((clean_txt(emp_code), clean_txt(name)))
    
    cursor.close()
    conn.close()
    return present_staff, late_staff, full_absent_staff

# ==========================================
# 3. INTERFACE RENDERING
# ==========================================
now_syria = datetime.now(SYRIA_TZ)
today_str = now_syria.strftime('%Y-%m-%d')
time_str = now_syria.strftime('%I:%M:%S %p')

st.title(TEXT_CONFIG["title_main"])
st.subheader(TEXT_CONFIG["title_sub"])
st.markdown(TEXT_CONFIG["lbl_date"].format(today_str, time_str))

if st.button(TEXT_CONFIG["btn_refresh"]):
    st.rerun()

try:
    present_staff, late_staff, full_absent_staff = load_attendance_data(today_str)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(TEXT_CONFIG["header_late"].format(len(late_staff)))
        if late_staff:
            for emp_code, name, time_in in late_staff:
                st.markdown(TEXT_CONFIG["late_row"].format(name, emp_code, time_in))
        else:
            st.success(TEXT_CONFIG["success_no_late"])
            
    with col2:
        st.markdown(TEXT_CONFIG["header_present"].format(len(present_staff)))
        if present_staff:
            for emp_code, name, time_in in present_staff:
                st.markdown(TEXT_CONFIG["present_row"].format(name, emp_code, time_in))
        else:
            st.info(TEXT_CONFIG["info_no_present"])
            
    with col3:
        st.markdown(TEXT_CONFIG["header_absent"].format(len(full_absent_staff)))
        if full_absent_staff:
            for emp_code, name in full_absent_staff:
                st.markdown(TEXT_CONFIG["absent_row"].format(name, emp_code))
        else:
            st.success(TEXT_CONFIG["success_no_absent"])
            
except Exception as e:
    st.error(TEXT_CONFIG["err_db"].format(str(e)))
    def sync_live_machine_deletions(local_cursor, neon_conn, neon_cursor):
    """Compares local active staff strings with cloud table entries and purges deleted records instantly along with constraints."""
    try:
        # 1. Pull all active employee codes from your local office machine database
        local_cursor.execute("SELECT emp_code FROM personnel_employee WHERE emp_code IS NOT NULL;")
        local_codes = [str(row[0]).strip() for row in local_cursor.fetchall() if row and row[0]]
        
        # Guard rail protect: If local database reads 0 records, do nothing to prevent accidental cloud wipes
        if not local_codes:
            return

        # 2. Query Neon Cloud to find out if any rows exist that are missing from your unpacked local array
        format_strings = ','.join('%s' for _ in local_codes)
        find_query = f"SELECT id, emp_code, first_name FROM personnel_employee WHERE emp_code NOT IN ({format_strings});"
        neon_cursor.execute(find_query, tuple(local_codes))
        stale_records = neon_cursor.fetchall()

        # 3. Safely clear privileges first, then delete the profile row completely from your cloud server
        for row in stale_records:
            emp_id, deleted_code, first_name = row
            print(f"🧹 Sync Engine Alert: User '{first_name}' (Code: {deleted_code}) was deleted from local system. Purging cloud records safely...")
            
            # Clear historical tracking records that are locking this specific user ID in the cloud database
            neon_cursor.execute("DELETE FROM acc_accprivilege WHERE employee_id = %s;", (emp_id,))
            neon_cursor.execute("DELETE FROM iclock_transaction WHERE emp_id = %s;", (emp_id,))
            
            # Now delete the main employee profile safely
            neon_cursor.execute("DELETE FROM personnel_employee WHERE id = %s;", (emp_id,))
            neon_conn.commit()
            print(f"✅ Successfully deleted Code {deleted_code} from Neon cluster cache.")
            
    except Exception as sync_err:
        print(f"⚠️ Live hardware deletion sync check skipped temporarily: {sync_err}")

