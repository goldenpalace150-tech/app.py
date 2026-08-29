import base64
from datetime import datetime, timedelta
import io
import unicodedata
import re
import zoneinfo
import zipfile
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import requests
import streamlit as strlit

# ==========================================
# 0. RTL ARABIC TEXT & VISUAL CONFIG
# ==========================================
TEXT_CONFIG = {
    "page_title": "حضور وانصراف القصر الذهبي",
    "title_main": "✨ شركة القصر الذهبي ✨",
    "search_placeholder": "🔍 ابحث باسم الموظف أو رقم الكود...",
    "header_all": "👥 كافة موظفي الشركة النشطين ({})",
    "header_present": "🟢 المتواجدون / الحضور ({})",
    "header_late": "⏰ الموظفون المتأخرون ({})",
    "header_checkout": "🏁 المنصرفون ({})",
    "header_leave": "🏖️ الموظفون في إجازة ({})",
    "header_absent": "❌ الغيابات ({})",
    "err_api": "خطأ في الاتصال بواجهة BioTime: {}",
}

strlit.set_page_config(
    page_title=TEXT_CONFIG["page_title"],
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ==========================================
# 1. CSS STYLING & ANIMATIONS
# ==========================================
strlit.markdown(
    """
    <style>
    header[data-testid="stHeader"] { display: none !important; }
    footer { display: none !important; }
    .stApp { direction: rtl; background-color: #f4f7f9; font-family: system-ui, -apple-system, sans-serif; }
    
    .block-container {
        padding-top: 15px !important;
        padding-bottom: 30px !important; 
        padding-left: 10px !important;
        padding-right: 10px !important;
        max-width: 100% !important;
    }

    .status-badge {
        display: flex;
        align-items: center;
        justify-content: center;
        background: linear-gradient(145deg, #ffffff, #f0f4f8);
        border: 1px solid #cbd5e1;
        padding: 8px 20px;
        border-radius: 30px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.04);
        margin: 0 auto 15px auto;
        width: fit-content;
        gap: 12px;
    }
    
    .animated-dish {
        width: 34px;
        height: 34px;
        object-fit: contain;
        transform-origin: center center;
        animation: rotate-360 5s linear infinite;
    }
    
    @keyframes rotate-360 {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }

    .status-indicator {
        display: flex;
        align-items: center;
        gap: 8px;
    }

    .blinking-dot {
        width: 10px;
        height: 10px;
        background-color: #22c55e;
        border-radius: 50%;
        display: inline-block;
        box-shadow: 0 0 0 0 rgba(34, 197, 94, 0.7);
        animation: pulse-green 1.5s infinite;
    }

    @keyframes pulse-green {
        0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(34, 197, 94, 0.7); }
        70% { transform: scale(1); box-shadow: 0 0 0 6px rgba(34, 197, 94, 0); }
        100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(34, 197, 94, 0); }
    }

    .online-text {
        font-size: 14px;
        font-weight: 800;
        color: #0f172a;
        letter-spacing: 0.5px;
    }

    div[data-testid="stColumn"] button {
        width: 100% !important;
        background: #ffffff !important;
        border-radius: 12px !important;
        padding: 12px 10px !important;
        text-align: center !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.04) !important;
        border: 1px solid #cbd5e1 !important;
        margin-bottom: 6px !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stColumn"] button p {
        font-size: 13px !important;
        color: #1e293b !important;
        font-weight: 700 !important;
        margin: 0 !important;
        white-space: pre-line !important;
        line-height: 1.4 !important;
    }

    .responsive-grid-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 15px;
        font-size: 13px;
        background-color: #ffffff;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 4px 6px rgba(0,0,0,0.02);
    }
    .responsive-grid-table .table-main-title-header {
        background: #0f172a;
        color: #ffffff !important;
        text-align: center;
        font-size: 14px;
        padding: 12px;
    }
    .responsive-grid-table th {
        background-color: #f8fafc;
        color: #475569;
        padding: 10px;
        border-bottom: 1px solid #e2e8f0;
        text-align: center;
    }
    .responsive-grid-table td { 
        padding: 10px; border-bottom: 1px solid #f1f5f9; text-align: center; font-weight: 500; color: #1e293b;
    }
    .badge-present { background-color: #dcfce7; color: #166534; padding: 4px 8px; border-radius: 6px; font-size: 11px; }
    .badge-late { background-color: #fef3c7; color: #9a3412; padding: 4px 8px; border-radius: 6px; font-size: 11px; }
    .badge-leave { background-color: #e0f2fe; color: #0369a1; padding: 4px 8px; border-radius: 6px; font-size: 11px; }
    .badge-absent { background-color: #fee2e2; color: #991b1b; padding: 4px 8px; border-radius: 6px; font-size: 11px; }
    </style>
""",
    unsafe_allow_html=True,
)

EXCLUDED_MANAGEMENT_CODES = ("40",)
SYRIA_TZ = zoneinfo.ZoneInfo("Asia/Damascus")

BASE_URL = strlit.secrets["biotime"]["base_url"].rstrip("/")
TOKEN_URL = strlit.secrets["biotime"]["token_url"]
EMAIL = strlit.secrets["biotime"]["email"]
PASSWORD = strlit.secrets["biotime"]["password"]
COMPANY = strlit.secrets["biotime"]["company"]

if "debug_logs" not in strlit.session_state:
  strlit.session_state["debug_logs"] = []
if "selected_view" not in strlit.session_state:
  strlit.session_state["selected_view"] = "present"


def clean_txt(raw_text):
  return (
      str(
          unicodedata.normalize("NFKC", str(raw_text))
          .replace("\u2066", "")
          .replace("\u2069", "")
          .strip()
      )
      if raw_text
      else ""
  )


@strlit.cache_data(ttl=300)
def get_auth_token():
  try:
    payload = {
        "username": EMAIL,
        "email": EMAIL,
        "password": PASSWORD,
        "company": COMPANY,
    }
    res = requests.post(TOKEN_URL, json=payload, timeout=15)

    if res.status_code in (200, 201):
      return res.json().get("token")
    else:
      strlit.error(f"BioTime Server Error (Code {res.status_code}): {res.text}")
      return None

  except requests.exceptions.Timeout:
    strlit.error("Connection Failed: BioTime server timed out.")
    return None
  except Exception as e:
    strlit.error(f"Connection Failed: {str(e)}")
    return None


def load_attendance_data_from_api(selected_date_str, selected_date_obj, is_today):
  token = get_auth_token()
  if not token:
    raise Exception(
        "تعذر المصادقة مع السيرفر. يرجى مراجعة تفاصيل الخطأ بالأعلى."
    )
  headers = {
      "Authorization": f"Token {token}",
      "Content-Type": "application/json",
  }

  devices = []
  try:
    for endpoint in ["/iclock/api/terminals/", "/iclock/api/devices/"]:
      dev_res = requests.get(
          f"{BASE_URL}{endpoint}", headers=headers, timeout=10
      )
      if dev_res.status_code == 200:
        d_data = dev_res.json()
        devices = (
            d_data.get("data", d_data)
            if isinstance(d_data, (dict, list))
            else []
        )
        break
  except Exception:
    pass

  terminal_map = {
      str(d.get("sn", "")): (
          d.get("alias") or d.get("terminal_name") or str(d.get("sn", ""))
      )
      for d in devices
      if d.get("sn")
  }

  all_employees = []
  try:
    emp_res = requests.get(
        f"{BASE_URL}/personnel/api/employees/?page_size=1000",
        headers=headers,
        timeout=15,
    )
    if emp_res.status_code == 200:
      all_employees = emp_res.json().get("data", [])
  except Exception:
    pass

  active_employees = {}
  for emp in all_employees:
    raw_code = str(emp.get("emp_code", "")).strip()
    cleaned_code = str(int(raw_code)) if raw_code.isdigit() else raw_code

    is_active = (
        str(emp.get("is_active", True)).lower() in ("true", "1", "yes")
    )
    emp_status = str(emp.get("status", "0")).upper()
    enable_att = (
        str(emp.get("enable_attendance", True)).lower() in ("true", "1", "yes")
    )

    if not is_active or emp_status in ("1", "2", "D") or not enable_att:
      continue

    if cleaned_code and cleaned_code not in EXCLUDED_MANAGEMENT_CODES:
      f_name = str(emp.get("first_name", "")).strip()
      l_name = str(emp.get("last_name", "")).strip()
      if f_name.lower() == "none":
        f_name = ""
      if l_name.lower() == "none":
        l_name = ""
      full_name = f"{f_name} {l_name}".strip()

      dept_data = emp.get("department", {})
      dept_name = (
          dept_data.get("dept_name")
          if isinstance(dept_data, dict)
          else str(emp.get("department", ""))
      )
      if not dept_name or dept_name.lower() == "none":
        dept_name = "غير محدد"

      active_employees[cleaned_code] = {
          "name": clean_txt(full_name if full_name else f"موظف {cleaned_code}"),
          "dept": clean_txt(dept_name),
      }

  leave_records = []
  try:
    leave_res = requests.get(
        f"{BASE_URL}/att/api/leave/?page_size=1000", headers=headers, timeout=10
    )
    if leave_res.status_code == 200:
      l_data = leave_res.json()
      leave_records = (
          l_data.get("data", l_data) if isinstance(l_data, (dict, list)) else []
      )
  except Exception:
    try:
      leave_res = requests.get(
          f"{BASE_URL}/iclock/api/leave/?page_size=1000",
          headers=headers,
          timeout=10,
      )
      if leave_res.status_code == 200:
        l_data = leave_res.json()
        leave_records = (
            l_data.get("data", l_data)
            if isinstance(l_data, (dict, list))
            else []
        )
    except Exception:
      pass

  on_leave_employees = {}
  for leave in leave_records:
    raw_code = str(
        leave.get("emp_code") or leave.get("employee_code") or ""
    ).strip()
    cleaned_code = str(int(raw_code)) if raw_code.isdigit() else raw_code
    start_t = (
        leave.get("start_time")
        or leave.get("start_date")
        or leave.get("start_datetime")
    )
    end_t = (
        leave.get("end_time")
        or leave.get("end_date")
        or leave.get("end_datetime")
    )

    leave_name = "إجازة"
    if "leave_type" in leave:
      if isinstance(leave["leave_type"], dict):
        leave_name = leave["leave_type"].get("leave_name", "إجازة")
      else:
        leave_name = str(leave["leave_type"])
    elif "leave_name" in leave:
      leave_name = leave["leave_name"]

    if cleaned_code and start_t and end_t:
      try:
        s_date = datetime.strptime(str(start_t)[:10], "%Y-%m-%d").date()
        e_date = datetime.strptime(str(end_t)[:10], "%Y-%m-%d").date()
        if s_date <= selected_date_obj <= e_date:
          on_leave_employees[cleaned_code] = clean_txt(leave_name)
      except Exception:
        pass

  prev_day = selected_date_obj.strftime("%Y-%m-%d") + " 00:00:00"
  next_day = (selected_date_obj + timedelta(days=1)).strftime(
      "%Y-%m-%d"
  ) + " 05:00:00"

  raw_logs = []
  try:
    logs_res = requests.get(
        f"{BASE_URL}/iclock/api/transactions/?start_time={prev_day}&end_time={next_day}&page_size=5000",
        headers=headers,
        timeout=15,
    )
    if logs_res.status_code == 200:
      raw_logs = logs_res.json().get("data", [])
  except Exception:
    pass

  emp_punches = {}
  for log in raw_logs:
    raw_code = str(log.get("emp_code", "")).strip()
    cleaned_code = str(int(raw_code)) if raw_code.isdigit() else raw_code
    if cleaned_code in active_employees and log.get("punch_time"):
      try:
        p_time = datetime.strptime(
            log.get("punch_time")[:19], "%Y-%m-%d %H:%M:%S"
        )
        dev_sn = str(log.get("terminal_sn", ""))
        dev_name = (
            log.get("terminal_alias")
            or log.get("terminal_name")
            or terminal_map.get(dev_sn, dev_sn or "جهاز رئيسي")
        )
        emp_punches.setdefault(cleaned_code, []).append((p_time, dev_name))
      except Exception:
        continue

  (
      present_staff,
      late_staff,
      absent_staff,
      checkout_staff,
      leave_staff,
      excel_rows,
  ) = ([], [], [], [], [], [])

  for code, emp_data in active_employees.items():
    name = emp_data["name"]
    dept = emp_data["dept"]
    punches = sorted(emp_punches.get(code, []), key=lambda x: x[0])
    filtered_punches = []

    for p_time, d_name in punches:
      if (
          not filtered_punches
          or abs((p_time - filtered_punches[-1][0]).total_seconds()) > 60
      ):
        filtered_punches.append((p_time, d_name))

    day_punches = [
        (p, d)
        for p, d in filtered_punches
        if p.date() == selected_date_obj and p.hour >= 5
    ]

    if not day_punches:
      if code in on_leave_employees:
        leave_reason = on_leave_employees[code]
        leave_staff.append((code, name, dept, leave_reason))
        excel_rows.append({
            "Employee ID": code,
            "First Name": name,
            "Department": dept,
            "Date": selected_date_str,
            "Clock In": "",
            "Clock Out": "",
            "Total WT": "",
            "Status": f"Leave - {leave_reason}",
        })
      else:
        absent_staff.append((code, name, dept))
        excel_rows.append({
            "Employee ID": code,
            "First Name": name,
            "Department": dept,
            "Date": selected_date_str,
            "Clock In": "",
            "Clock Out": "",
            "Total WT": "",
            "Status": "Absence(A)",
        })
      continue

    first_p, first_dev = day_punches[0]
    is_late = first_p.hour > 9 or (first_p.hour == 9 and first_p.minute > 15)

    next_morning = [
        (p, d)
        for p, d in filtered_punches
        if p.date() == selected_date_obj + timedelta(days=1) and p.hour < 5
    ]
    punch_count = (
        2 if (len(day_punches) % 2 != 0 and next_morning) else len(day_punches)
    )

    last_p = None
    last_dev = first_dev

    if punch_count % 2 == 0:
      last_p, last_dev = (
          next_morning[-1]
          if (len(day_punches) % 2 != 0 and next_morning)
          else day_punches[-1]
      )
    elif not is_today and len(day_punches) > 1:
      last_p, last_dev = day_punches[-1]

    total_wt_str = ""
    if last_p and first_p:
      diff = last_p - first_p
      hours, remainder = divmod(int(diff.total_seconds()), 3600)
      minutes = remainder // 60
      total_wt_str = f"{hours:02d}:{minutes:02d}"

    status_str = "Late(LT)" if is_late else "Present(P)"

    if is_late:
      late_staff.append(
          (code, name, dept, first_p.strftime("%I:%M %p"), first_dev)
      )

    if is_today:
      if punch_count % 2 != 0:
        present_staff.append(
            (code, name, dept, first_p.strftime("%I:%M %p"), first_dev)
        )
        excel_rows.append({
            "Employee ID": code,
            "First Name": name,
            "Department": dept,
            "Date": selected_date_str,
            "Clock In": first_p.strftime("%H:%M"),
            "Clock Out": "",
            "Total WT": "",
            "Status": status_str,
        })
      else:
        last_p_real, last_dev_real = (
            next_morning[-1]
            if (len(day_punches) % 2 != 0 and next_morning)
            else day_punches[-1]
        )
        checkout_staff.append((
            code,
            name,
            dept,
            last_p_real.strftime("%I:%M %p"),
            last_dev_real,
        ))
        excel_rows.append({
            "Employee ID": code,
            "First Name": name,
            "Department": dept,
            "Date": selected_date_str,
            "Clock In": first_p.strftime("%H:%M"),
            "Clock Out": last_p_real.strftime("%H:%M"),
            "Total WT": total_wt_str,
            "Status": status_str,
        })
    else:
      if last_p:
        checkout_staff.append(
            (code, name, dept, last_p.strftime("%I:%M %p"), last_dev)
        )
      else:
        present_staff.append(
            (code, name, dept, first_p.strftime("%I:%M %p"), first_dev)
        )

      excel_rows.append({
          "Employee ID": code,
          "First Name": name,
          "Department": dept,
          "Date": selected_date_str,
          "Clock In": first_p.strftime("%H:%M"),
          "Clock Out": last_p.strftime("%H:%M") if last_p else "",
          "Total WT": total_wt_str,
          "Status": status_str,
      })

  absent_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else 999)
  present_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else 999)
  late_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else 999)
  leave_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else 999)
  checkout_staff.sort(key=lambda x: int(x[0]) if x[0].isdigit() else 999)

  return (
      active_employees,
      present_staff,
      late_staff,
      absent_staff,
      checkout_staff,
      leave_staff,
      devices,
      excel_rows,
  )


# ==========================================
# 3. INTERFACE RENDERING
# ==========================================
now_syria = datetime.now(SYRIA_TZ)
today_str = now_syria.strftime("%Y-%m-%d")

dish_img_tag = ""
try:
  with open("image_632b3d.jpg", "rb") as img_file:
    encoded_string = base64.b64encode(img_file.read()).decode()
    dish_img_tag = f'<img src="data:image/jpeg;base64,{encoded_string}" class="animated-dish" />'
except Exception:
  dish_img_tag = '<div class="animated-dish" style="font-size: 24px;">📡</div>'

c_date, c_ref = strlit.columns(2)
with c_date:
  selected_date_obj_input = strlit.date_input(
      "", value=now_syria.date(), label_visibility="collapsed"
  )
  selected_date_str = selected_date_obj_input.strftime("%Y-%m-%d")
with c_ref:
  if strlit.button("🔄 تحديث البيانات", use_container_width=True):
    strlit.cache_data.clear()
    strlit.rerun()

is_today = selected_date_str == today_str

if "last_selected_date" not in strlit.session_state:
  strlit.session_state["last_selected_date"] = selected_date_str

if strlit.session_state["last_selected_date"] != selected_date_str:
  strlit.session_state["last_selected_date"] = selected_date_str
  strlit.session_state["selected_view"] = "present" if is_today else "all"

if is_today:
  strlit.markdown(
      f"""
        <div class="status-badge">
            {dish_img_tag}
            <div class="status-indicator">
                <span class="blinking-dot"></span>
                <span class="online-text">Online (مباشر)</span>
            </div>
        </div>
        """,
      unsafe_allow_html=True,
  )
else:
  strlit.markdown(
      f"""
        <div class="status-badge" style="border-color: #94a3b8; background: #e2e8f0;">
            {dish_img_tag}
            <div class="status-indicator">
                <span style="font-weight: 800; color: #475569; font-size: 14px;">أرشيف تاريخي ({selected_date_str})</span>
            </div>
        </div>
        """,
      unsafe_allow_html=True,
  )

try:
  act, pre, lat, abs_s, chk, lev, devices, exc = load_attendance_data_from_api(
      selected_date_str, selected_date_obj_input, is_today
  )

  # 📥 UPLOAD TEMPLATE & FILL ATTENDANCE VALUES OR GENERATE DEFAULT REPORT
  col_gen, col_up = strlit.columns(2)

  with col_gen:
    # Standard Generated Attendance Report
    df_excel = pd.DataFrame(exc)
    output = io.BytesIO()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Attendance Report"

    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3"),
    )

    headers = [
        "Employee ID",
        "First Name",
        "Department",
        "Date",
        "Clock In",
        "Clock Out",
        "Total WT",
        "Status",
    ]
    ws.append(headers)
    ws.row_dimensions[1].height = 24

    for col_idx in range(1, len(headers) + 1):
      cell = ws.cell(row=1, column=col_idx)
      cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
      cell.fill = PatternFill(
          start_color="1F4E78", end_color="1F4E78", fill_type="solid"
      )
      cell.alignment = Alignment(horizontal="center", vertical="center")
      cell.border = thin_border

    for idx, row_data in enumerate(exc, 2):
      ws.row_dimensions[idx].height = 20
      ws.append([
          row_data["Employee ID"],
          row_data["First Name"],
          row_data["Department"],
          row_data["Date"],
          row_data["Clock In"],
          row_data["Clock Out"],
          row_data["Total WT"],
          row_data["Status"],
      ])

      status_val = str(row_data["Status"])
      row_fill = None
      row_font_color = "000000"

      if "Leave" in status_val or "L" in status_val:
        row_fill = PatternFill(
            start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
        )
        row_font_color = "002060"
      elif "Absence" in status_val or "A" in status_val:
        row_fill = PatternFill(
            start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
        )
        row_font_color = "9C0006"

      for col_idx in range(1, 9):
        cell = ws.cell(row=idx, column=col_idx)

        if row_fill:
          cell.fill = row_fill
          cell.font = Font(
              name="Calibri", size=11, bold=True, color=row_font_color
          )
        else:
          cell.font = Font(name="Calibri", size=11)

        if col_idx == 8 and not row_fill:
          if "Late" in status_val or "LT" in status_val:
            cell.font = Font(
                name="Calibri", size=11, bold=True, color="9C0006"
            )
            cell.fill = PatternFill(
                start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
            )
          elif "Present" in status_val or "P" in status_val:
            cell.font = Font(
                name="Calibri", size=11, bold=True, color="006100"
            )
            cell.fill = PatternFill(
                start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
            )

        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for col in ws.columns:
      max_len = 0
      col_letter = get_column_letter(col[0].column)
      for cell in col:
        if cell.value:
          max_len = max(max_len, len(str(cell.value)))
      ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

    wb.save(output)
    excel_data = output.getvalue()

    strlit.download_button(
        label="📥 تحميل تقرير اليوم (افتراضي)",
        data=excel_data,
        file_name=f"Daily_Attendance_Report_{selected_date_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

  with col_up:
    uploaded_template = strlit.file_uploader(
        "📂 رفع جدول الدوام الشهري وتعبئته تلقائياً",
        type=["xlsx", "xlsm"],
        label_visibility="collapsed",
        key="monthly_attendance_template",
    )

    def normalize_id(value):
      """Normalize an employee ID without changing its identity."""
      if value is None:
        return ""

      # Excel may give a numeric ID as 1.0 when the source cell is numeric.
      if isinstance(value, float) and value.is_integer():
        value = int(value)

      raw = str(value).strip()
      if not raw:
        return ""

      if raw.endswith(".0") and raw[:-2].isdigit():
        raw = raw[:-2]

      return raw

    def detect_month_sheet(workbook):
      preferred = [
          sheet_name
          for sheet_name in workbook.sheetnames
          if any(
              key in sheet_name
              for key in [
                  "مجموع ساعات الدوام",
                  "دوام",
                  "ساعات",
                  "حضور",
                  "August",
                  "اغسطس",
                  "أغسطس",
              ]
          )
      ]
      return workbook[
          preferred[0] if preferred else workbook.sheetnames[0]
      ]

    def detect_date_columns(ws):
      """The attendance template has its dates in row 1."""
      date_columns = {}

      for column in range(1, ws.max_column + 1):
        value = ws.cell(row=1, column=column).value
        parsed = None

        if isinstance(value, datetime):
          parsed = value.date()
        elif hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
          parsed = value
        elif isinstance(value, str):
          raw = value.strip()
          for fmt in (
              "%Y-%m-%d",
              "%d/%m/%Y",
              "%d-%m-%Y",
              "%d/%m/%y",
              "%d-%m-%y",
          ):
            try:
              parsed = datetime.strptime(raw, fmt).date()
              break
            except ValueError:
              continue

        if parsed:
          date_columns[column] = parsed

      return date_columns

    def time_to_excel_value(value):
      """Convert HH:MM to a real Excel time fraction."""
      if value is None or str(value).strip() == "":
        return None

      raw = str(value).strip()
      parts = raw.split(":")
      if len(parts) < 2:
        return None

      try:
        hours = int(parts[0])
        minutes = int(parts[1])
        seconds = int(parts[2]) if len(parts) > 2 else 0
        return (hours * 3600 + minutes * 60 + seconds) / 86400
      except (TypeError, ValueError):
        return None

    def excel_col_letter(column_number):
      return get_column_letter(column_number)

    def patch_cell_value_xml(xml_text, row_number, column_number, value):
      """Change one cell in worksheet XML with a minimal text patch.

      This preserves the original XML structure, namespaces, metadata, tables,
      printer settings, cloud extensions, and any unsupported Excel features.
      """
      coordinate = f"{get_column_letter(column_number)}{row_number}"

      # Match either a self-closing cell or a normal cell element while preserving
      # every attribute except the value/type attributes we intentionally change.
      pattern = re.compile(
          rf'<c\b(?=[^>]*\br="{re.escape(coordinate)}")([^>]*)/>'
          rf'|<c\b(?=[^>]*\br="{re.escape(coordinate)}")([^>]*)>.*?</c>',
          re.DOTALL,
      )

      match = pattern.search(xml_text)
      if not match:
        raise RuntimeError(f"Excel cell {coordinate} was not found in worksheet XML")

      attrs = match.group(1) if match.group(1) is not None else match.group(2)

      # Remove an existing type attribute. We put back only the type we need.
      attrs = re.sub(r'\s+t="[^"]*"', '', attrs)

      if value is None or value == "":
        replacement = f"<c{attrs}/>"
      elif isinstance(value, str):
        safe_text = (
            value.replace("&", "&amp;")
                 .replace("<", "&lt;")
                 .replace(">", "&gt;")
                 .replace('"', "&quot;")
        )
        replacement = f'<c{attrs} t="inlineStr"><is><t>{safe_text}</t></is></c>'
      else:
        replacement = f"<c{attrs}><v>{value}</v></c>"

      return xml_text[:match.start()] + replacement + xml_text[match.end():]

    def find_sheet_xml_path(zip_file, sheet_name):
      """Resolve a workbook sheet name to its worksheet XML path."""
      main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      pkg_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"

      workbook_xml = zip_file.read("xl/workbook.xml").decode("utf-8")
      rels_xml = zip_file.read("xl/_rels/workbook.xml.rels").decode("utf-8")

      workbook_root = ET.fromstring(workbook_xml)
      rel_root = ET.fromstring(rels_xml)

      rel_targets = {}
      for rel in rel_root.findall(f"{{{pkg_rel_ns}}}Relationship"):
        rel_targets[rel.get("Id")] = rel.get("Target")

      for sheet in workbook_root.findall(f"{{{main_ns}}}sheets/{{{main_ns}}}sheet"):
        if sheet.get("name") == sheet_name:
          rel_id = sheet.get(f"{{{rel_ns}}}id")
          target = rel_targets.get(rel_id)
          if not target:
            raise RuntimeError(f"Could not resolve worksheet relationship for {sheet_name}")
          target = target.lstrip("/")
          if not target.startswith("xl/"):
            target = "xl/" + target
          return target

      raise RuntimeError(f"Worksheet not found: {sheet_name}")

    def export_template_preserving_package(original_bytes, sheet_name, cell_updates):
      """Patch attendance values into the original OOXML package.

      Unlike openpyxl.save(), this does not rebuild the workbook. It copies the
      original ZIP package byte-for-byte except for the specific attendance cells
      that need changing. This is designed to avoid damaging modern Excel metadata
      and workbook features that can be lost during a full openpyxl rewrite.
      """
      input_buffer = io.BytesIO(original_bytes)
      output_buffer = io.BytesIO()

      with zipfile.ZipFile(input_buffer, "r") as zin:
        sheet_xml_path = find_sheet_xml_path(zin, sheet_name)
        sheet_xml_text = zin.read(sheet_xml_path).decode("utf-8")

        for update in cell_updates:
          sheet_xml_text = patch_cell_value_xml(
              sheet_xml_text,
              update["row"],
              update["column"],
              update["value"],
          )

        patched_sheet_xml = sheet_xml_text.encode("utf-8")

        with zipfile.ZipFile(output_buffer, "w") as zout:
          for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == sheet_xml_path:
              data = patched_sheet_xml
            zout.writestr(item, data)

      output_buffer.seek(0)
      return output_buffer.getvalue()

    @strlit.cache_data(ttl=300, show_spinner=False)
    def fetch_month_date(attendance_date):
      date_str = attendance_date.strftime("%Y-%m-%d")
      is_date_today = attendance_date == now_syria.date()
      return load_attendance_data_from_api(
          date_str,
          attendance_date,
          is_date_today,
      )

    if uploaded_template is not None:
      try:
        # Keep the exact original XLSX/XLSM bytes. The final export will patch
        # only the attendance cells into this original package.
        original_template_bytes = uploaded_template.getvalue()
        keep_vba = uploaded_template.name.lower().endswith(".xlsm")

        template_wb = openpyxl.load_workbook(
            io.BytesIO(original_template_bytes),
            keep_vba=keep_vba,
            data_only=False,
        )
        ws_target = detect_month_sheet(template_wb)

        date_columns = detect_date_columns(ws_target)
        if not date_columns:
          strlit.error("لم يتم العثور على أعمدة التواريخ في الصف الأول من ملف الدوام.")
          raise RuntimeError("No monthly date columns detected")

        # IMPORTANT:
        # Excel Column B is the BioTime ID.
        # Excel Column A is NOT used for employee matching.
        # Employee names are NOT used for matching.
        excel_employees = []
        for row in range(2, ws_target.max_row + 1):
          biotime_id = normalize_id(ws_target.cell(row=row, column=2).value)
          employee_name = clean_txt(ws_target.cell(row=row, column=3).value)

          if not biotime_id and not employee_name:
            continue

          excel_employees.append(
              {
                  "row": row,
                  "biotime_id": biotime_id,
                  "name": employee_name,
              }
          )

        if not excel_employees:
          strlit.error(
              "لم يتم العثور على موظفين في الملف. تأكد من أن BioTime ID موجود في العمود B."
          )
          raise RuntimeError("No employee rows detected")

        # Fetch BioTime for every date in the Excel file up to today.
        all_attendance = {}
        employee_catalog = {}
        progress = strlit.progress(
            0,
            text="جاري تحميل بيانات الدوام لجميع التواريخ...",
        )

               total_dates = len(dates_list)

        for index, (date_column, attendance_date) in enumerate(
            dates_list,
            start=1,
        ):
          (
              active_employees,
              _present,
              _late,
              _absent,
              _checkout,
              _leave,
              _devices,
              excel_rows_for_date,
          ) = fetch_month_date(attendance_date)

          for employee_id, employee_data in active_employees.items():
            normalized_employee_id = normalize_id(employee_id)
            employee_catalog[normalized_employee_id] = clean_txt(
                employee_data.get("name", "")
            )

          date_map = {}
          for item in excel_rows_for_date:
            employee_id = normalize_id(item.get("Employee ID", ""))
            if employee_id:
              date_map[employee_id] = item
          all_attendance[attendance_date] = date_map

          progress.progress(
              index / max(total_dates, 1),
              text=(
                  f"جاري معالجة {attendance_date.strftime('%d/%m/%Y')} "
                  f"({index}/{total_dates})"
              ),
          )

        progress.empty()

        # EXACT MATCH ONLY:
        # Excel Column B (BioTime ID) == app/BioTime ID (emp_code).
        # No name fallback and no fuzzy matching.
        employee_matches = []
        unmatched_employees = []
        duplicate_excel_ids = set()
        seen_excel_ids = set()

        for excel_employee in excel_employees:
          excel_id = excel_employee["biotime_id"]

          if not excel_id:
            unmatched_employees.append(
                {
                    "row": excel_employee["row"],
                    "biotime_id": "",
                    "excel_name": excel_employee["name"],
                    "reason": "BioTime ID is empty",
                }
            )
            continue

          if excel_id in seen_excel_ids:
            duplicate_excel_ids.add(excel_id)
          seen_excel_ids.add(excel_id)

          if excel_id in employee_catalog:
            employee_matches.append(
                {
                    "row": excel_employee["row"],
                    "excel_name": excel_employee["name"],
                    "employee_id": excel_id,
                    "api_name": employee_catalog[excel_id],
                    "score": 100,
                    "match_type": "Exact ID",
                }
            )
          else:
            unmatched_employees.append(
                {
                    "row": excel_employee["row"],
                    "biotime_id": excel_id,
                    "excel_name": excel_employee["name"],
                    "reason": "BioTime ID not found in app/BioTime ID list",
                }
            )

        preview_rows = [
            {
                "اسم الملف": match["excel_name"],
                "BioTime ID (Excel B)": match["employee_id"],
                "ID (App/BioTime)": match["employee_id"],
                "المطابقة": "Exact ID",
            }
            for match in employee_matches
        ]

        strlit.success(
            f"تم العثور على {len(sorted_dates)} تاريخ و{len(employee_matches)} موظف مطابق بالـ ID فقط."
        )
        if duplicate_excel_ids:
          strlit.warning(
              "يوجد BioTime ID مكرر في ملف Excel: "
              + ", ".join(sorted(duplicate_excel_ids))
          )
        if unmatched_employees:
          strlit.warning(
              f"يوجد {len(unmatched_employees)} موظف غير مطابق. لن يتم استيراد أي حضور لهم."
          )

        if preview_rows:
          strlit.dataframe(
              pd.DataFrame(preview_rows),
              use_container_width=True,
              hide_index=True,
          )

        if unmatched_employees:
          with strlit.expander("⚠️ الموظفون غير المطابقين", expanded=False):
            strlit.dataframe(
                pd.DataFrame(unmatched_employees),
                use_container_width=True,
                hide_index=True,
            )

        if strlit.button(
            "⚙️ تشغيل تعبئة جميع التواريخ",
            use_container_width=True,
            key="run_monthly_attendance",
        ):
          filled_cells = 0
          cell_updates = []
          import_log = []

          # Process every date column. Only exact BioTime ID matches are eligible.
          for match in employee_matches:
            employee_id = match["employee_id"]
            excel_row = match["row"]

            for date_column, attendance_date in date_columns.items():
              # Future dates stay unchanged/blank.
              if attendance_date > now_syria.date():
                continue

              attendance = all_attendance.get(
                  attendance_date,
                  {},
              ).get(employee_id)

              if attendance is None:
                cell_value = "A"
                status = "Absence(A)"
                clock_in = ""
                clock_out = ""
                total_work = ""
              else:
                status = str(attendance.get("Status", ""))
                clock_in = str(attendance.get("Clock In", "") or "")
                clock_out = str(attendance.get("Clock Out", "") or "")
                total_work = str(attendance.get("Total WT", "") or "")

                if "Leave" in status:
                  cell_value = "L"
                elif "Absence" in status:
                  cell_value = "A"
                else:
                  excel_time = time_to_excel_value(total_work)
                  cell_value = excel_time if excel_time is not None else None

              cell_updates.append(
                  {
                      "row": excel_row,
                      "column": date_column,
                      "value": cell_value,
                  }
              )
              filled_cells += 1

              import_log.append(
                  [
                      employee_id,
                      match["excel_name"],
                      match["api_name"],
                      "Exact ID",
                      attendance_date,
                      clock_in,
                      clock_out,
                      total_work,
                      status,
                  ]
              )

          try:
            export_bytes = export_template_preserving_package(
                original_template_bytes,
                ws_target.title,
                cell_updates,
            )
          except Exception as export_error:
            raise RuntimeError(
                "تعذر إنشاء ملف Excel النهائي مع الحفاظ على بنية الملف الأصلية: "
                + str(export_error)
            ) from export_error

          strlit.session_state["monthly_attendance_export"] = export_bytes
          output_extension = ".xlsm" if keep_vba else ".xlsx"
          strlit.session_state["monthly_attendance_filename"] = (
              "Attendance_Completed_"
              f"{sorted_dates[0].strftime('%Y_%m_%d')}"
              "_to_"
              f"{sorted_dates[-1].strftime('%Y_%m_%d')}"
              f"{output_extension}"
          )
          strlit.session_state["monthly_attendance_mime"] = (
              "application/vnd.ms-excel.sheet.macroEnabled.12"
              if keep_vba
              else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          )
          strlit.session_state["monthly_attendance_import_summary"] = {
              "matched": len(employee_matches),
              "unmatched": len(unmatched_employees),
              "filled": filled_cells,
          }

          strlit.success(
              f"✅ تمت تعبئة جميع التواريخ حتى {now_syria.date().strftime('%d/%m/%Y')} "
              f"({filled_cells} خانة). تم تصدير الملف مع الحفاظ على بنية Excel الأصلية."
          )

        if "monthly_attendance_export" in strlit.session_state:
          summary = strlit.session_state.get("monthly_attendance_import_summary", {})
          if summary:
            strlit.info(
                f"مطابق: {summary.get('matched', 0)} | "
                f"غير مطابق: {summary.get('unmatched', 0)} | "
                f"خانات الدوام المعبأة: {summary.get('filled', 0)}"
            )

          strlit.download_button(
              label="📥 تحميل ملف الدوام الجاهز",
              data=strlit.session_state["monthly_attendance_export"],
              file_name=strlit.session_state["monthly_attendance_filename"],
              mime=strlit.session_state.get(
                  "monthly_attendance_mime",
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              ),
              use_container_width=True,
              key="download_monthly_attendance",
          )

      except Exception as template_error:
        strlit.error(
            f"تعذر معالجة ملف الدوام: {template_error}"
        )
  if is_today:
    if strlit.button(
        f"👥 كافة موظفي الشركة النشطين ({len(act)})", use_container_width=True
    ):
      strlit.session_state["selected_view"] = "all"

    col_p, col_l = strlit.columns(2)
    with col_p:
      if strlit.button(
          f"🟢 المتواجدون ({len(pre)})", use_container_width=True
      ):
        strlit.session_state["selected_view"] = "present"
    with col_l:
      if strlit.button(f"⏰ المتأخرون ({len(lat)})", use_container_width=True):
        strlit.session_state["selected_view"] = "late"

    col_c, col_a = strlit.columns(2)
    with col_c:
      if strlit.button(
          f"🏁 المنصرفون ({len(chk)})", use_container_width=True
      ):
        strlit.session_state["selected_view"] = "checkout"
    with col_a:
      if strlit.button(
          f"❌ الغيابات ({len(abs_s)})", use_container_width=True
      ):
        strlit.session_state["selected_view"] = "absent"

    col_lv, col_dummy = strlit.columns(2)
    with col_lv:
      if strlit.button(
          f"🏖️ الإجازات ({len(lev)})", use_container_width=True
      ):
        strlit.session_state["selected_view"] = "leave"

  # 🖨️ DEVICES EXPANDER (Checks 30-min offline limit)
  with strlit.expander("🖨️ أجهزة الحضور والانصراف المرتبطة", expanded=False):
    if devices:
      dev_rows = []
      for d in devices:
        d_name = (
            d.get("alias")
            or d.get("terminal_name")
            or d.get("sn", "جهاز غير محدد")
        )
        d_sn = d.get("sn", "N/A")
        d_ip = d.get("ip_address", "غير متوفر")

        last_activity = d.get("last_activity")
        status_badge = "<span class='badge-absent'>غير متصل 🔴</span>"
        if last_activity:
          try:
            last_act_dt = datetime.strptime(
                last_activity[:19], "%Y-%m-%d %H:%M:%S"
            )
            if (
                datetime.now().replace(tzinfo=None) - last_act_dt
            ).total_seconds() < 1800:
              status_badge = "<span class='badge-present'>متصل 🟢</span>"
          except Exception:
            pass

        dev_rows.append(
            f"<tr><td>{d_name}</td><td>{d_sn}</td><td>{d_ip}</td><td>{status_badge}</td></tr>"
        )
      strlit.markdown(
          f'<table class="responsive-grid-table"><tr><th>اسم'
          " الجهاز</th><th>الرقم التسلسلي (SN)</th><th>عنوان"
          f' IP</th><th>الحالة</th></tr>{"".join(dev_rows)}</table>',
          unsafe_allow_html=True,
      )

  search_query = (
      strlit.text_input(
          "",
          placeholder=TEXT_CONFIG["search_placeholder"],
          label_visibility="collapsed",
      )
      .strip()
      .lower()
  )
  match = (
      lambda c, n: (
          search_query in str(c).lower() or search_query in str(n).lower()
      )
      if search_query
      else True
  )

  view = strlit.session_state["selected_view"]

  if view == "all":
    rows = [
        f"<tr><td>{c}</td><td>{d_data['name']}</td><td>{d_data['dept']}</td><td><span"
        " class='badge-present'>نشط</span></td></tr>"
        for c, d_data in act.items()
        if match(c, d_data["name"])
    ]
    strlit.markdown(
        '<table class="responsive-grid-table"><tr><th colspan="4"'
        f' class="table-main-title-header">{TEXT_CONFIG["header_all"].format(len(act))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>القسم</th><th>الحالة</th></tr>{"".join(rows)}</table>',
        unsafe_allow_html=True,
    )

  elif view == "present":
    if pre:
      rows = [
          f"<tr><td>{c}</td><td>{n}</td><td>{dpt}</td><td>{t}</td><td>{d}</td></tr>"
          for c, n, dpt, t, d in pre
          if match(c, n)
      ]
      strlit.markdown(
          '<table class="responsive-grid-table"><tr><th colspan="5"'
          f' class="table-main-title-header">{TEXT_CONFIG["header_present"].format(len(pre))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>القسم</th><th>الدخول</th><th>جهاز'
          f' البصمة</th></tr>{"".join(rows)}</table>',
          unsafe_allow_html=True,
      )

  elif view == "late":
    if lat:
      rows = [
          f"<tr><td>{c}</td><td>{n}</td><td>{dpt}</td><td>{t}</td><td><span"
          f" class='badge-late'>متأخر</span></td><td>{d}</td></tr>"
          for c, n, dpt, t, d in lat
          if match(c, n)
      ]
      strlit.markdown(
          '<table class="responsive-grid-table"><tr><th colspan="6"'
          f' class="table-main-title-header">{TEXT_CONFIG["header_late"].format(len(lat))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>القسم</th><th>الدخول</th><th>الحالة</th><th>جهاز'
          f' البصمة</th></tr>{"".join(rows)}</table>',
          unsafe_allow_html=True,
      )

  elif view == "checkout":
    if chk:
      rows = [
          f"<tr><td>{c}</td><td>{n}</td><td>{dpt}</td><td>{t}</td><td>{d}</td></tr>"
          for c, n, dpt, t, d in chk
          if match(c, n)
      ]
      strlit.markdown(
          '<table class="responsive-grid-table"><tr><th colspan="5"'
          f' class="table-main-title-header">{TEXT_CONFIG["header_checkout"].format(len(chk))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>القسم</th><th>الانصراف</th><th>جهاز'
          f' البصمة</th></tr>{"".join(rows)}</table>',
          unsafe_allow_html=True,
      )

  elif view == "leave":
    if lev:
      rows = [
          f"<tr><td>{c}</td><td>{n}</td><td>{dpt}</td><td><span"
          f" class='badge-leave'>{r}</span></td></tr>"
          for c, n, dpt, r in lev
          if match(c, n)
      ]
      strlit.markdown(
          '<table class="responsive-grid-table"><tr><th colspan="4"'
          f' class="table-main-title-header">{TEXT_CONFIG["header_leave"].format(len(lev))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>القسم</th><th>نوع'
          f' الإجازة</th></tr>{"".join(rows)}</table>',
          unsafe_allow_html=True,
      )

  elif view == "absent":
    if abs_s:
      rows = [
          f"<tr><td>{c}</td><td>{n}</td><td>{dpt}</td><td><span"
          " class='badge-absent'>غياب</span></td></tr>"
          for c, n, dpt in abs_s
          if match(c, n)
      ]
      strlit.markdown(
          '<table class="responsive-grid-table"><tr><th colspan="4"'
          f' class="table-main-title-header">{TEXT_CONFIG["header_absent"].format(len(abs_s))}</th></tr><tr><th>الكود</th><th>الاسم</th><th>القسم</th><th>الحالة</th></tr>{"".join(rows)}</table>',
          unsafe_allow_html=True,
      )

except Exception as e:
  strlit.error(TEXT_CONFIG["err_api"].format(str(e)))
