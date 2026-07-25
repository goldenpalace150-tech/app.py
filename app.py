import streamlit as st
import psycopg2
import unicodedata
from datetime import datetime
import zoneinfo

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
    "lbl_date": "📅 التاريخ الحالي في سوريا: **{}**  │  ⏰ الوقت الحالي: **{}**",
    "lbl_picker": "📅 اختر التاريخ المراد عرض بياناته:",
    "btn_refresh": "🔄 تحديث البيانات الحية الآن",
    "header_late": "⏰ المتأخرون اليوم ({}) – دخول بعد 09:15 صباحاً",
    "late_row": "🔸 **{}** (كود: {}) ── وقت الدخول: {}",
    "success_no_late": "🎉 لا يوجد متأخرين في هذا التاريخ!",
    "header_absent": "❌ غائبون أو نسوا تسجيل الحضور ({})",
    "absent_row": "🔹 **{}** (كود: {})",
    "success_no_absent": "🎉 لا يوجد غيابات في هذا التاريخ!",
    "header_present": "🟢 الموظفون المتواجدون حالياً في العمل ({})",
    "present_row": "🔸 **{}** (كود: {}) ── وقت الدخول: {}",
    "info_no_present": "لا يوجد موظفين منتظمين متواجدين حالياً.",
    "err_db": "خطأ في الاتصال بقاعدة البيانات السحابية: {}"
}

st.set_page_config(page_title=TEXT_CONFIG["page_title"], page_icon="📊", layout="wide")
st.markdown(TEXT_CONFIG["style_align"], unsafe_allow_html=True)

# Exclude management codes from visual dashboard lists
EXCLUDED_MANAGEMENT_CODES = ("40", "10", "20")
mgmt_codes_str = ",".join(f"'{code}'" for code in EXCLUDED_MANAGEMENT_CODES)
DATABASE_URL = st.secrets["NEON_DATABASE_URL"]
SYRIA_TZ = zoneinfo.ZoneInfo("Asia/Damascus")

def clean_txt(raw_text):
    if not raw_text: return ""
    return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip())

def load_attendance_data(selected_date_str):
    conn = psycopg2.connect(DATABASE_URL)
    cursor = conn.cursor()
    
    # 1. Query present and late staff using raw data comparisons (bypassing timezone math bugs)
    query1 = f"""
        SELECT e.emp_code, e.first_name, 
               MIN(t.punch_time) as first_punch,
               COUNT(t.id) as punch_count
        FROM personnel_employee e 
        JOIN iclock_transaction t ON e.id = t.emp_id
        WHERE t.punch_time::date = '{selected_date_str}' 
          AND e.status = 1
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
        
        # Check hours using the raw local punch time directly
        if first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15):
            late_staff.append((emp_code, clean_name, time_in_clean))
        
        if punch_count % 2 != 0:
            present_staff.append((emp_code, clean_name, time_in_clean))
            
    # 2. Query full absent staff using matching raw date formats
    query0 = f"""
        SELECT DISTINCT e.emp_code, e.first_name FROM personnel_employee e
        WHERE e.id NOT IN (SELECT DISTINCT emp_id FROM iclock_transaction WHERE punch_time::date = '{selected_date_str}')
          AND e.status = 1
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
current_today = now_syria.date()
time_str = now_syria.strftime('%I:%M:%S %p')

st.title(TEXT_CONFIG["title_main"])
st.subheader(TEXT_CONFIG["title_sub"])
st.markdown(TEXT_CONFIG["lbl_date"].format(current_today.strftime('%Y-%m-%d'), time_str))

# Interactive local system date picker
selected_date = st.date_input(TEXT_CONFIG["lbl_picker"], value=current_today)
selected_date_str = selected_date.strftime('%Y-%m-%d')

if st.button(TEXT_CONFIG["btn_refresh"]):
    st.rerun()

try:
    present_staff, late_staff, full_absent_staff = load_attendance_data(selected_date_str)
    
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
