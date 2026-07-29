import streamlit as st
import requests
import unicodedata
import pandas as pd
import io
from datetime import datetime, timedelta
import zoneinfo

from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 0. RTL ARABIC TEXT & VISUAL CONFIG
# ==========================================
TEXT_CONFIG = {
    "page_title": "حضور وانصراف القصر الذهبي",
    "title_main": "✨ شركة القصر الذهبي ✨",
    "lbl_date": "📅 التاريخ المختار للتقرير: **{}**  │  ⏰ الوقت الحالي في سوريا: **{}**",
    "btn_refresh": "🔄 تحديث البيانات الحية الآن",
    "lbl_pick_date": "📅 اختر تاريخ عرض التقرير:",
    "btn_download_excel": "📥 تحميل تقرير الحضور الشامل كملف Excel النمطي",
    
    "header_all": "👥 قائمة كافة موظفي الشركة النشطين ({})",
    "header_present": "🟢 قائمة الموظفين المتواجدون حالياً ({})",
    "header_late": "⏰ قائمة الموظفين المتأخرين اليوم ({})",
    "header_checkout": "🏁 قائمة الموظفين المنصرفين اليوم ({})",
    "header_absent": "❌ قائمة الغيابات الكاملة اليوم ({})",
    
    "late_row": "🔸 **{}** (كود: {}) ── وقت الدخول: {}",
    "absent_row": "🔹 **{}** (كود: {})",
    "present_row": "🔸 **{}** (كود: {}) ── وقت الدخول: {}",
    "checkout_row": "✅ **{}** (كود: {}) ── وقت الانصراف: {}",
    "all_row": "👤 **{}** (كود: {})",
    
    "err_api": "خطأ في الاتصال بواجهة BioTime السحابية: {}"
}

st.set_page_config(page_title=TEXT_CONFIG["page_title"], page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; }
    .reportview-container .main .block-container { direction: RTL; text-align: right; }
    h1, h2, h3, h4, p, span, li, div { text-align: right !important; direction: RTL !important; line-height: 1.6 !important; }
    
    div.stButton > button {
        width: 100% !important;
        background-color: #ffffff !important;
        color: #1e293b !important;
        border: 1px solid #eef2f5 !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.06) !important;
        border-radius: 8px !important;
        padding: 15px !important;
        text-align: right !important;
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        margin-bottom: 2px !important;
    }
    div.stButton > button:hover {
        border-color: #cbd5e1 !important;
        background-color: #f8fafc !important;
    }
    
    .miracle-banner {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        color: #ffffff !important;
        border-radius: 12px;
        padding: 25px;
        text-align: center !important;
        box-shadow: 0 10px 15px -3px rgba(59, 130, 246, 0.2);
        margin-top: 15px;
        margin-bottom: 25px;
        direction: rtl;
    }
    .miracle-title { font-size: 22px; font-weight: bold; margin-bottom: 10px; color: #ffffff !important; }
    .miracle-text { font-size: 15px; opacity: 0.95; line-height: 1.8 !important; color: #f0fdf4 !important; }
    
    .list-wrapper-box {
        background-color: #f8fafc;
        border: 1px dashed #cbd5e1;
        border-radius: 6px;
        padding: 15px;
        margin-top: 5px;
        margin-bottom: 15px;
        direction: rtl;
    }
    </style>
""", unsafe_allow_html=True)

EXCLUDED_MANAGEMENT_CODES = ("40",)
EXCLUDED_RESIGNED_CODES = ("34",) 

SYRIA_TZ = zoneinfo.ZoneInfo("Asia/Damascus")

BASE_URL = st.secrets["biotime"]["base_url"].rstrip('/')
TOKEN_URL = st.secrets["biotime"]["token_url"]
EMAIL = st.secrets["biotime"]["email"]
PASSWORD = st.secrets["biotime"]["password"]
COMPANY = st.secrets["biotime"]["company"]

if "debug_logs" not in st.session_state:
    st.session_state["debug_logs"] = []

if "selected_view" not in st.session_state:
    st.session_state["selected_view"] = "present"

def log_debug(message):
    st.session_state["debug_logs"].append(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

def clean_txt(raw_text):
    if not raw_text: return ""
    return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip())
@st.cache_data(ttl=300)
def get_auth_token():
    payload = {"email": EMAIL, "password": PASSWORD, "company": COMPANY}
    try:
        response = requests.post(TOKEN_URL, json=payload, timeout=10)
        if response.status_code == 200 or response.status_code == 201:
            return response.json().get("token")
        return None
    except Exception:
        return None

def load_attendance_data_from_api(selected_date_str, selected_date_obj):
    token = get_auth_token()
    if not token:
        raise Exception("تفاصيل رمز المصادقة (Token) مفقودة أو غير صالحة.")
        
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    st.session_state["debug_logs"] = []
    
    emp_url = f"{BASE_URL}/personnel/api/employees/?page_size=1000"
    all_employees = []
    try:
        emp_res = requests.get(emp_url, headers=headers, timeout=15)
        if emp_res.status_code == 200:
            all_employees = emp_res.json().get("data", [])
    except Exception as e:
        log_debug(f"Employee Request Error: {str(e)}")

    active_employees = {}
    for emp in all_employees:
        raw_code = str(emp.get("emp_code", "")).strip()
        if not raw_code:
            continue
        cleaned_code = str(int(raw_code)) if raw_code.isdigit() else raw_code
        if cleaned_code not in EXCLUDED_MANAGEMENT_CODES and cleaned_code not in EXCLUDED_RESIGNED_CODES:
            first_name = emp.get("first_name", "") or ""
            last_name = emp.get("last_name", "") or ""
            full_name = f"{first_name} {last_name}".strip()
            active_employees[cleaned_code] = clean_txt(full_name if full_name else f"User {cleaned_code}")

    prev_day_obj = selected_date_obj - timedelta(days=1)
    next_day_obj = selected_date_obj + timedelta(days=1)
    
    start_window_str = prev_day_obj.strftime('%Y-%m-%d') + " 00:00:00"
    end_window_str = next_day_obj.strftime('%Y-%m-%d') + " 05:00:00"
    
    logs_url = f"{BASE_URL}/iclock/api/transactions/?start_time={start_window_str}&end_time={end_window_str}&page_size=5000"
    raw_logs = []
    try:
        logs_res = requests.get(logs_url, headers=headers, timeout=15)
        if logs_res.status_code == 200:
            raw_logs = logs_res.json().get("data", [])
    except Exception as e:
        log_debug(f"Logs Request Error: {str(e)}")

    emp_punches = {}
    for log in raw_logs:
        raw_code = str(log.get("emp_code", "")).strip()
        if not raw_code:
            continue
        cleaned_code = str(int(raw_code)) if raw_code.isdigit() else raw_code
        if cleaned_code in active_employees:
            punch_time_str = log.get("punch_time", "")
            if punch_time_str:
                try:
                    p_time = datetime.strptime(punch_time_str[:19], "%Y-%m-%d %H:%M:%S")
                    if cleaned_code not in emp_punches:
                        emp_punches[cleaned_code] = []
                    emp_punches[cleaned_code].append(p_time)
                except Exception:
                    continue

    present_staff = []      
    late_staff = []         
    full_absent_staff = []  
    checkout_staff = []     
    excel_rows = [] 

    for code, name in active_employees.items():
        user_all_punches = sorted(emp_punches.get(code, []))
        
        # Isolation and cleaning of active transactions array
        cleaned_current_day_punches = []
        day_raw_punches = [p for p in user_all_punches if p.date() == selected_date_obj]
        
        for p in day_raw_punches:
            # 🛡️ THE CROSSOVER OVERWRITE FIX: Deletes early-morning isolated logs before 5:00 AM 
            # if they belong to yesterday's night shift
            if p.hour < 5:
                yesterday_raw = [x for x in user_all_punches if x.date() == prev_day_obj]
                if not yesterday_raw:
                    continue
                else:
                    yesterday_clean = []
                    for yp in yesterday_raw:
                        if yp.hour < 5:
                            day_before_p = [x for x in user_all_punches if x.date() == (prev_day_obj - timedelta(days=1))]
                            if day_before_p and len(day_before_p) % 2 != 0:
                                continue
                        yesterday_clean.append(yp)
                    if yesterday_clean and len(yesterday_clean) % 2 != 0:
                        continue
            cleaned_current_day_punches.append(p)

        if not cleaned_current_day_punches:
            full_absent_staff.append((code, name))
            excel_rows.append({
                "Employee ID": code, "First Name": name, "Date": selected_date_str,
                "Clock In": "", "Clock Out": "", "Total WT": "", "Status": "Absence(A)"
            })
            continue

        first_punch = cleaned_current_day_punches[0]
        clock_in_str = first_punch.strftime('%H:%M')
        
        is_late = first_punch.hour > 9 or (first_punch.hour == 9 and first_punch.minute > 15)
        status_label = "Late(LT)" if is_late else "Present(P)"
        
        if is_late:
            late_staff.append((code, name, first_punch.strftime('%I:%M %p')))

        early_morning_punches_next_day = [p for p in user_all_punches if p.date() == next_day_obj and p.hour < 5]
        if len(cleaned_current_day_punches) % 2 != 0 and early_morning_punches_next_day:
            last_punch = early_morning_punches_next_day[0]
            punch_count = 2
        else:
            last_punch = cleaned_current_day_punches[-1]
            punch_count = len(cleaned_current_day_punches)

        if punch_count % 2 != 0:
            present_staff.append((code, name, first_punch.strftime('%I:%M %p')))
            excel_rows.append({
                "Employee ID": code, "First Name": name, "Date": selected_date_str,
                "Clock In": clock_in_str, "Clock Out": "", "Total WT": "", "Status": status_label
            })
        else:
            clock_out_str = last_punch.strftime('%H:%M')
            checkout_staff.append((code, name, last_punch.strftime('%I:%M %p')))
            
            time_diff = last_punch - first_punch
            total_seconds = int(time_diff.total_seconds())
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            total_wt_str = f"{hours:02d}:{minutes:02d}"
            
            excel_rows.append({
                "Employee ID": code, "First Name": name, "Date": selected_date_str,
                "Clock In": clock_in_str, "Clock Out": clock_out_str, "Total WT": total_wt_str, "Status": status_label
            })

    full_absent_staff.sort(key=lambda val: int(val) if str(val).isdigit() else 999)
    present_staff.sort(key=lambda val: int(val) if str(val).isdigit() else 999)
    late_staff.sort(key=lambda val: int(val) if str(val).isdigit() else 999)
    checkout_staff.sort(key=lambda val: int(val) if str(val).isdigit() else 999)
    excel_rows.sort(key=lambda row_item: int(row_item["Employee ID"]) if str(row_item["Employee ID"]).isdigit() else 999)
    
    return active_employees, present_staff, late_staff, full_absent_staff, checkout_staff, excel_rows
# ==========================================
# 3. INTERFACE RENDERING & CONTROLS
# ==========================================
now_syria = datetime.now(SYRIA_TZ)
time_str = now_syria.strftime('%I:%M:%S %p')

selected_date = st.date_input(TEXT_CONFIG["lbl_pick_date"], value=now_syria.date())
selected_date_str = selected_date.strftime('%Y-%m-%d')

formatted_excel_date = selected_date.strftime('%B %d %Y')
generated_on_timestamp = now_syria.strftime('%a %b %d %Y %H:%M:%S')

st.title(TEXT_CONFIG["title_main"])
st.markdown(TEXT_CONFIG["lbl_date"].format(selected_date_str, time_str))

btn_col1, btn_col2 = st.columns(2)
with btn_col1:
    if st.button(TEXT_CONFIG["btn_refresh"], use_container_width=True):
        st.cache_data.clear()
        st.rerun()

try:
    active_employees, present_staff, late_staff, full_absent_staff, checkout_staff, excel_rows = load_attendance_data_from_api(selected_date_str, selected_date)
    
    excel_buffer = io.BytesIO()
    df_grid_data = pd.DataFrame(excel_rows)
    
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df_grid_data.to_excel(writer, sheet_name="Attendance Report", index=False, startrow=5)
        workbook = writer.book
        worksheet = writer.sheets["Attendance Report"]
        
        worksheet["A1"] = "Daily Attendance Report(Basic Report)"
        worksheet["A1"].font = Font(name="Arial", size=16, bold=True)
        worksheet["A1"].alignment = Alignment(horizontal="center")
        worksheet.merge_cells("A1:G1")
        
        worksheet["A2"] = formatted_excel_date
        worksheet["A2"].font = Font(name="Arial", size=12, bold=False)
        worksheet["A2"].alignment = Alignment(horizontal="center")
        worksheet.merge_cells("A2:G2")
        
        worksheet["A3"] = "Company: Golden Palace"
        worksheet["A3"].font = Font(name="Arial", size=11, bold=False)
        worksheet["A3"].alignment = Alignment(horizontal="left")
        worksheet.merge_cells("A3:C3")
        
        worksheet["D3"] = f"Generated On: {generated_on_timestamp}"
        worksheet["D3"].font = Font(name="Arial", size=11, bold=False)
        worksheet["D3"].alignment = Alignment(horizontal="right")
        worksheet.merge_cells("D3:G3")
        
        worksheet["A5"] = "Department: Department 1"
        worksheet["A5"].font = Font(name="Arial", size=11, bold=True)
        worksheet["A5"].alignment = Alignment(horizontal="left")
        worksheet.merge_cells("A5:C5")
        
        thin_border_side = Side(border_style="thin", color="000000")
        grid_border_format = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
        
        for col_idx in range(1, 8):
            cell = worksheet.cell(row=6, column=col_idx)
            cell.font = Font(name="Arial", size=11, bold=True)
            cell.border = grid_border_format
            cell.alignment = Alignment(horizontal="center")
            
        for row in worksheet.iter_rows(min_row=7, max_row=worksheet.max_row, min_col=1, max_col=7):
            for cell in row:
                cell.font = Font(name="Arial", size=10)
                cell.border = grid_border_format
                cell.alignment = Alignment(horizontal="center" if cell.column != 2 else "left")
                
        for col_cells in worksheet.columns:
            max_len = 0
            first_cell = col_cells[0]
            col_letter = get_column_letter(first_cell.column)
            
            for cell in col_cells:
                if cell.row <= 5:
                    continue
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    with btn_col2:
        st.download_button(
            label=TEXT_CONFIG["btn_download_excel"],
            data=excel_buffer.getvalue(),
            file_name=f"Daily_Attendance_Report_{selected_date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    total_emp = len(active_employees)
    p_count = len(present_staff)
    l_count = len(late_staff)
    c_count = len(checkout_staff)
    a_count = len(full_absent_staff)
    
    st.write("### 📊 اضغط على أي بطاقة لعرض أسماء الموظفين أسفلها مباشرة")
    current_view = st.session_state["selected_view"]
    is_today = (selected_date == now_syria.date())
    
    def show_miracle_message():
        st.markdown(f"""
            <div class="miracle-banner">
                <div class="miracle-title">✨ هذا اليوم مؤرشف ومغلق بالكامل ✨</div>
                <div class="miracle-text">
                    لقد تم إغلاق وتأمين كشوفات يوم <b>{selected_date_str}</b> بنجاح داخل النظام. <br>
                    يمكنك تحميل التقرير الشامل والمطابق للهيكل التنظيمي عبر زر الـ <b>Excel</b> بالأعلى لمراجعة ساعات العمل الإجمالية.
                </div>
            </div>
        """, unsafe_allow_html=True)

    # 1. Total Registered Card Button
    if st.button(f"👥 إجمالي عدد موظفي الشركة النشطين ── {total_emp if is_today else 0}"):
        st.session_state["selected_view"] = "all"
        st.rerun()
    if current_view == "all":
        if is_today:
            st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
            st.write(f"### {TEXT_CONFIG['header_all'].format(total_emp)}")
            for code, name in sorted(active_employees.items(), key=lambda x: int(x) if str(x).isdigit() else 999):
                st.markdown(TEXT_CONFIG["all_row"].format(name, code))
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            show_miracle_message()
        
    # 2. Present Count Card Button
    if st.button(f"🟢 الموظفون المتواجدون حالياً في العمل ── {p_count if is_today else 0}"):
        st.session_state["selected_view"] = "present"
        st.rerun()
    if current_view == "present":
        if is_today:
            st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
            st.write(f"### {TEXT_CONFIG['header_present'].format(p_count)}")
            if present_staff:
                for code, name, time_in in present_staff:
                    st.markdown(TEXT_CONFIG["present_row"].format(name, code, time_in))
            else:
                st.info("لا يوجد موظفين متواجدين حالياً داخل المنشأة.")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            show_miracle_message()
        
    # 3. Late Count Card Button
    if st.button(f"⏰ الموظفون المتأخرون اليوم ── {l_count if is_today else 0}"):
        st.session_state["selected_view"] = "late"
        st.rerun()
    if current_view == "late":
        if is_today:
            st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
            st.write(f"### {TEXT_CONFIG['header_late'].format(l_count)}")
            if late_staff:
                for code, name, time_in in late_staff:
                    st.markdown(TEXT_CONFIG["late_row"].format(name, code, time_in))
            else:
                st.success("🎉 لا يوجد متأخرين اليوم!")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            show_miracle_message()
        
    # 4. Checked-Out Card Button
    if st.button(f"✅ الموظفون الذين غادروا وانصرفوا ── {c_count if is_today else 0}"):
        st.session_state["selected_view"] = "checkout"
        st.rerun()
    if current_view == "checkout":
        if is_today:
            st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
            st.write(f"### {TEXT_CONFIG['header_checkout'].format(c_count)}")
            if checkout_staff:
                for code, name, time_out in checkout_staff:
                    st.markdown(TEXT_CONFIG["checkout_row"].format(name, code, time_out))
            else:
                st.info("لا توجد عمليات انصراف مسجلة حتى الآن.")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            show_miracle_message()
        
    # 5. Absent Card Button
    if st.button(f"❌ الموظفون الغائبون بالكامل اليوم ── {a_count if is_today else 0}"):
        st.session_state["selected_view"] = "absent"
        st.rerun()
    if current_view == "absent":
        if is_today:
            st.markdown('<div class="list-wrapper-box">', unsafe_allow_html=True)
            st.write(f"### {TEXT_CONFIG['header_absent'].format(a_count)}")
            if full_absent_staff:
                for code, name in full_absent_staff:
                    st.markdown(TEXT_CONFIG["absent_row"].format(name, code))
            else:
                st.success("🎉 لا يوجد غيابات اليوم!")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            show_miracle_message()

except Exception as e:
    st.error(TEXT_CONFIG["err_api"].format(str(e)))
    if st.checkbox("عرض سجلات الأخطاء البرمجية التشغيلية (Debug Logs)"):
        for log in st.session_state.get("debug_logs", []):
            st.text(log)
