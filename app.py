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
APP_VERSION = "BIO-ATTENDANCE-FAST-SINGLE-PUNCH-TOTALS-2026-08-30"

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
# A lone punch at or after this hour is treated as an OUT punch.
SINGLE_PUNCH_OUT_HOUR = 14
# BioTime timetable used by the current attendance setup. The Basic Report
# confirms that Actual WT for a single punch is the worked overlap with
# 09:00-19:00, capped at 10 hours.
SINGLE_PUNCH_SHIFT_START_HOUR = 9
SINGLE_PUNCH_SHIFT_END_HOUR = 19
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


def calculate_single_punch_actual_wt(punch_time, is_out_punch):
  """Return BioTime-equivalent Actual WT for one missing punch.

  Current company timetable is 09:00-19:00 (10 hours). For a missing OUT,
  Actual WT is from the recorded IN up to 19:00. For a missing IN, it is from
  09:00 up to the recorded OUT. Early/late excess outside the timetable is
  capped exactly as BioTime's Basic Report does.
  """
  shift_start = punch_time.replace(
      hour=SINGLE_PUNCH_SHIFT_START_HOUR,
      minute=0,
      second=0,
      microsecond=0,
  )
  shift_end = punch_time.replace(
      hour=SINGLE_PUNCH_SHIFT_END_HOUR,
      minute=0,
      second=0,
      microsecond=0,
  )
  max_seconds = int((shift_end - shift_start).total_seconds())

  if is_out_punch:
    effective_out = min(max(punch_time, shift_start), shift_end)
    worked_seconds = int((effective_out - shift_start).total_seconds())
  else:
    effective_in = max(min(punch_time, shift_end), shift_start)
    worked_seconds = int((shift_end - effective_in).total_seconds())

  worked_seconds = max(0, min(max_seconds, worked_seconds))
  hours, remainder = divmod(worked_seconds, 3600)
  minutes = remainder // 60
  return f"{hours:02d}:{minutes:02d}"


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

    next_morning = [
        (p, d)
        for p, d in filtered_punches
        if p.date() == selected_date_obj + timedelta(days=1) and p.hour < 5
    ]

    # For exactly one punch, classify an early punch as IN and an afternoon/
    # evening punch as OUT. Calculate the same schedule-aware Actual WT shown
    # by BioTime's Basic Report, without requiring a second upload or report API.
    single_punch_only = len(day_punches) == 1 and not next_morning
    single_punch_is_out = (
        single_punch_only and first_p.hour >= SINGLE_PUNCH_OUT_HOUR
    )

    is_late = (
        not single_punch_is_out
        and (first_p.hour > 9 or (first_p.hour == 9 and first_p.minute > 15))
    )

    if single_punch_only:
      clock_in_value = "" if single_punch_is_out else first_p.strftime("%H:%M")
      clock_out_value = first_p.strftime("%H:%M") if single_punch_is_out else ""
      single_punch_actual_wt = calculate_single_punch_actual_wt(
          first_p,
          single_punch_is_out,
      )

      if single_punch_is_out:
        checkout_staff.append(
            (code, name, dept, first_p.strftime("%I:%M %p"), first_dev)
        )
        status_str = "Present(P) / Missing IN"
      else:
        present_staff.append(
            (code, name, dept, first_p.strftime("%I:%M %p"), first_dev)
        )
        status_str = (
            "Late(LT) / Missing OUT" if is_late else "Present(P) / Missing OUT"
        )
        if is_late:
          late_staff.append(
              (code, name, dept, first_p.strftime("%I:%M %p"), first_dev)
          )

      excel_rows.append({
          "Employee ID": code,
          "First Name": name,
          "Department": dept,
          "Date": selected_date_str,
          "Clock In": clock_in_value,
          "Clock Out": clock_out_value,
          "Actual WT": single_punch_actual_wt,
          "Calculated WT": single_punch_actual_wt,
          "Total WT": single_punch_actual_wt,
          "Status": status_str,
      })
      continue

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
    # ONE upload box only: the monthly attendance Excel template.
    # BioTime calculated work hours are fetched automatically from BioTime.
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

    def normalize_report_key(value):
      """Normalize report JSON field names for tolerant matching."""
      return re.sub(r"[^a-z0-9]", "", str(value or "").lower())

    def flatten_report_row(value, prefix=""):
      """Flatten nested report JSON while retaining short and full keys."""
      flat = {}

      if isinstance(value, dict):
        for key, child in value.items():
          normalized_key = normalize_report_key(key)
          full_key = f"{prefix}{normalized_key}" if prefix else normalized_key

          if isinstance(child, (dict, list)):
            flat.update(flatten_report_row(child, full_key))
          else:
            if normalized_key and normalized_key not in flat:
              flat[normalized_key] = child
            if full_key:
              flat[full_key] = child

      elif isinstance(value, list):
        for item in value:
          flat.update(flatten_report_row(item, prefix))

      return flat

    def extract_report_rows(payload):
      """Extract attendance rows from BioTime report JSON.

      BioTime releases return report data in several shapes. Some are flat DRF
      lists, while others group daily rows beneath an employee object. Preserve
      scalar parent fields while walking nested lists so employee/date context is
      not lost before normalization.
      """
      rows = []
      seen = set()

      row_hint_keys = {
          "attdate",
          "attendancedate",
          "workdate",
          "date",
          "clockin",
          "clockout",
          "checkin",
          "checkout",
          "actualwt",
          "actualworked",
          "actualworktime",
          "totalwt",
          "totalworked",
          "totalworktime",
      }

      def walk(value, inherited=None):
        inherited = dict(inherited or {})

        if isinstance(value, list):
          for item in value:
            walk(item, inherited)
          return

        if not isinstance(value, dict):
          return

        scalar_context = dict(inherited)
        nested_items = []

        for key, child in value.items():
          if isinstance(child, (dict, list)):
            nested_items.append((key, child))
            # Employee/report group metadata is often a nested object sitting
            # beside the actual daily rows. Carry its scalar fields into sibling
            # row context using prefixed names such as employee_emp_code.
            if isinstance(child, dict):
              for child_key, child_value in child.items():
                if not isinstance(child_value, (dict, list)):
                  scalar_context.setdefault(
                      f"{key}_{child_key}",
                      child_value,
                  )
          else:
            scalar_context.setdefault(key, child)

        normalized_keys = {normalize_report_key(key) for key in value.keys()}
        has_row_hint = bool(normalized_keys & row_hint_keys)

        if has_row_hint:
          merged = dict(inherited)
          merged.update(value)
          # Avoid adding the same object twice when several envelope paths point
          # to the identical dictionary.
          signature = id(value)
          if signature not in seen:
            seen.add(signature)
            rows.append(merged)

        for _key, child in nested_items:
          walk(child, scalar_context)

      walk(payload)

      # Flat report lists occasionally contain fields whose names are unknown to
      # row_hint_keys. Fall back to the conventional envelopes if needed.
      if rows:
        return rows

      if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]
      if isinstance(payload, dict):
        for key in ("data", "results", "rows", "items"):
          candidate = payload.get(key)
          if isinstance(candidate, list):
            return [item for item in candidate if isinstance(item, dict)]
      return []

    def first_report_value(flat_row, candidate_keys):
      """Return the first non-empty report field, preferring exact keys."""
      normalized_candidates = [
          normalize_report_key(candidate_key) for candidate_key in candidate_keys
      ]

      # Exact keys first. flatten_report_row keeps short child keys, so this
      # avoids mistaking e.g. total_work_time for the more specific work_time.
      for normalized_candidate in normalized_candidates:
        if not normalized_candidate:
          continue
        value = flat_row.get(normalized_candidate)
        if value not in (None, ""):
          return value, normalized_candidate

      # Fallback for unusual nested serializers that expose only prefixed keys.
      for normalized_candidate in normalized_candidates:
        if not normalized_candidate:
          continue
        for key, value in flat_row.items():
          if value in (None, ""):
            continue
          if key.endswith(normalized_candidate):
            return value, key
      return "", ""

    def parse_report_date(value, fallback_date=None):
      if isinstance(value, datetime):
        return value.date()
      if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
        return value
      if value not in (None, ""):
        raw = str(value).strip()

        # BioTime versions may serialize dates as YYYY-MM-DD, full timestamps,
        # or ISO-8601 strings with a T/Z timezone marker.
        try:
          return datetime.fromisoformat(raw.replace("Z", "+00:00")).date()
        except ValueError:
          pass

        for candidate, fmt in (
            (raw[:10], "%Y-%m-%d"),
            (raw[:19], "%Y-%m-%d %H:%M:%S"),
            (raw[:10], "%d/%m/%Y"),
            (raw[:10], "%d-%m-%Y"),
            (raw[:8], "%d/%m/%y"),
            (raw[:8], "%d-%m-%y"),
        ):
          try:
            return datetime.strptime(candidate, fmt).date()
          except ValueError:
            continue
      return fallback_date

    def duration_to_hhmm(value, source_key=""):
      """Normalize a BioTime calculated duration to HH:MM."""
      if value is None or str(value).strip() == "":
        return ""

      if isinstance(value, timedelta):
        total_seconds = int(value.total_seconds())
      elif isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
        key = normalize_report_key(source_key)
        if "second" in key:
          total_seconds = round(number)
        elif "minute" in key:
          total_seconds = round(number * 60)
        elif 0 <= number <= 1 and not number.is_integer():
          # Some report serializers return an Excel/day fraction.
          total_seconds = round(number * 86400)
        elif 0 <= number <= 24:
          total_seconds = round(number * 3600)
        elif 0 <= number <= 1440:
          total_seconds = round(number * 60)
        else:
          total_seconds = round(number)
      else:
        raw = str(value).strip()

        day_match = re.match(
            r"^(?P<days>\d+)\s+day[s]?,\s*(?P<hours>\d+):(?P<minutes>\d{1,2})(?::(?P<seconds>\d{1,2}))?$",
            raw,
            re.IGNORECASE,
        )
        if day_match:
          total_seconds = (
              int(day_match.group("days")) * 86400
              + int(day_match.group("hours")) * 3600
              + int(day_match.group("minutes")) * 60
              + int(day_match.group("seconds") or 0)
          )
        else:
          time_match = re.match(
              r"^(?P<hours>\d+):(?P<minutes>\d{1,2})(?::(?P<seconds>\d{1,2}))?$",
              raw,
          )
          if time_match:
            total_seconds = (
                int(time_match.group("hours")) * 3600
                + int(time_match.group("minutes")) * 60
                + int(time_match.group("seconds") or 0)
            )
          else:
            try:
              number = float(raw.replace(",", "."))
            except ValueError:
              return ""
            total_seconds = round(number * 3600)

      if total_seconds < 0:
        return ""

      hours, remainder = divmod(total_seconds, 3600)
      minutes = remainder // 60
      return f"{hours:02d}:{minutes:02d}"

    def normalize_clock_value(value):
      """Normalize a report clock value to HH:MM without inventing a punch."""
      if value in (None, ""):
        return ""
      if isinstance(value, datetime):
        return value.strftime("%H:%M")
      if hasattr(value, "hour") and hasattr(value, "minute"):
        try:
          return f"{int(value.hour):02d}:{int(value.minute):02d}"
        except Exception:
          pass

      raw = str(value).strip()
      if not raw:
        return ""

      for fmt in (
          "%Y-%m-%d %H:%M:%S",
          "%Y-%m-%d %H:%M",
          "%d-%m-%Y %H:%M:%S",
          "%d/%m/%Y %H:%M:%S",
          "%H:%M:%S",
          "%H:%M",
          "%I:%M %p",
      ):
        try:
          return datetime.strptime(raw[:19], fmt).strftime("%H:%M")
        except ValueError:
          continue

      # Keep an already-recognisable HH:MM prefix.
      match = re.search(r"(?<!\d)(\d{1,2}):(\d{2})(?::\d{2})?", raw)
      if match:
        try:
          hour = int(match.group(1))
          minute = int(match.group(2))
          if 0 <= hour <= 23 and 0 <= minute <= 59:
            return f"{hour:02d}:{minute:02d}"
        except ValueError:
          pass
      return raw

    def normalize_calculated_report_row(
        report_row,
        start_date,
        end_date,
        employee_internal_id_map=None,
    ):
      """Convert one BioTime attendance-report row to the app attendance shape.

      The BioTime report APIs are not fully consistent between releases. Some
      versions return emp_code directly; others return an internal employee id or
      a nested employee object. The same is true for worked-time field names.
      """
      flat = flatten_report_row(report_row)
      employee_internal_id_map = employee_internal_id_map or {}

      employee_value, _employee_key = first_report_value(
          flat,
          (
              "emp_code",
              "employee_emp_code",
              "employee_code",
              "employee_code_code",
              "employeecode",
              "empcode",
              "staff_code",
              "staffcode",
              "employee_number",
              "employee_no",
          ),
      )
      employee_id = normalize_id(employee_value)

      if not employee_id:
        internal_value, _internal_key = first_report_value(
            flat,
            (
                "employee_id",
                "employeeid",
                "emp_id",
                "empid",
                "employee",
                "emp",
            ),
        )
        internal_id = normalize_id(internal_value)
        employee_id = employee_internal_id_map.get(internal_id, "")

      if not employee_id:
        return None

      date_value, _date_key = first_report_value(
          flat,
          (
              "att_date",
              "attendance_date",
              "attendance_day",
              "work_date",
              "workday",
              "attdate",
              "transaction_date",
              "date",
          ),
      )
      attendance_date = parse_report_date(date_value)
      if attendance_date is None or not (start_date <= attendance_date <= end_date):
        return None

      clock_in, _clock_in_key = first_report_value(
          flat,
          (
              "clock_in",
              "check_in",
              "checkin",
              "first_in",
              "first_punch",
              "first_punch_time",
              "checkin_time",
              "in_time",
              "punch_in",
              "first_clock_in",
          ),
      )
      clock_out, _clock_out_key = first_report_value(
          flat,
          (
              "clock_out",
              "check_out",
              "checkout",
              "last_out",
              "last_punch",
              "last_punch_time",
              "checkout_time",
              "out_time",
              "punch_out",
              "last_clock_out",
          ),
      )

      # Prefer BioTime's actual worked duration. This is the value that recovers
      # a missing IN/OUT according to the employee's timetable. Total duration is
      # used only when the report does not expose an actual-worked field.
      actual_value, actual_key = first_report_value(
          flat,
          (
              "actual_wt",
              "actual_work_time",
              "actual_working_time",
              "actual_worked_time",
              "actual_worked_hours",
              "actual_worked",
              "actual_work",
              "actual_time",
              "actual_duration",
              "actual_work_duration",
              "worked_time",
              "worked_hours",
              "worked_hrs",
              "working_hours",
              "working_hrs",
              "work_time",
              "work_hours",
              "work_hrs",
              "worked",
              "work_duration",
              "worked_duration",
              "actual_worked_minutes",
              "actual_worked_seconds",
              "worked_minutes",
              "worked_seconds",
          ),
      )
      actual_work_time = duration_to_hhmm(actual_value, actual_key)

      total_value, total_key = first_report_value(
          flat,
          (
              "total_wt",
              "total_worked",
              "total_work_time",
              "total_worked_time",
              "total_worked_hours",
              "total_duration",
              "total_hours",
              "total_hrs",
              "total_time",
              "duration",
          ),
      )
      total_work_time = duration_to_hhmm(total_value, total_key)
      calculated_work_time = actual_work_time or total_work_time

      clock_in = normalize_clock_value(clock_in)
      clock_out = normalize_clock_value(clock_out)

      # Ignore report metadata rows that do not describe attendance.
      if not calculated_work_time and not clock_in and not clock_out:
        return None

      if clock_in and not clock_out:
        status = "Present(P) / Missing OUT"
      elif clock_out and not clock_in:
        status = "Present(P) / Missing IN"
      elif calculated_work_time:
        status = "Present(P)"
      else:
        status = ""

      return {
          "Employee ID": employee_id,
          "Attendance Date": attendance_date,
          "Date": attendance_date.strftime("%Y-%m-%d"),
          "Clock In": clock_in,
          "Clock Out": clock_out,
          "Actual WT": actual_work_time,
          "Report Total WT": total_work_time,
          "Calculated WT": calculated_work_time,
          "Total WT": calculated_work_time,
          "Status": status,
      }

    def report_record_score(record, expected_single_punch=False):
      """Prefer the report row most useful for single-punch recovery."""
      score = 0
      if expected_single_punch:
        score += 100
      if record.get("Actual WT"):
        score += 50
      if record.get("Report Total WT"):
        score += 30
      if record.get("Calculated WT"):
        score += 20
      if bool(record.get("Clock In")) != bool(record.get("Clock Out")):
        score += 10
      return score

    def fetch_report_pages(session, endpoint, query, request_kwargs):
      """Fetch one BioTime report query and follow DRF pagination."""
      payloads = []
      next_url = endpoint
      next_params = dict(query)
      seen_urls = set()

      for _page in range(20):
        request_key = (next_url, tuple(sorted((next_params or {}).items())))
        if request_key in seen_urls:
          break
        seen_urls.add(request_key)

        try:
          response = session.get(
              next_url,
              params=next_params,
              timeout=20,
              **request_kwargs,
          )
        except requests.RequestException:
          return [], None

        if response.status_code != 200:
          return [], response.status_code

        try:
          payload = response.json()
        except ValueError:
          return [], response.status_code

        payloads.append(payload)

        if not isinstance(payload, dict):
          break
        next_link = payload.get("next")
        if not next_link:
          break
        next_link = str(next_link)
        if next_link.startswith("http://") or next_link.startswith("https://"):
          next_url = next_link
        elif next_link.startswith("/"):
          next_url = BASE_URL + next_link
        else:
          next_url = BASE_URL + "/" + next_link.lstrip("/")
        next_params = None

      return payloads, 200

    @strlit.cache_data(ttl=300, show_spinner=False)
    def fetch_biotime_calculated_range(
        start_date,
        end_date,
        employee_internal_id_map=None,
        expected_single_punch_keys=None,
    ):
      """Fetch BioTime-calculated hours for a date range in a few bulk calls.

      Only BioTime report endpoints are used here. The Basic Report file does not
      need to be uploaded by the user.
      """
      start_text = start_date.strftime("%Y-%m-%d")
      end_text = end_date.strftime("%Y-%m-%d")
      employee_internal_id_map = employee_internal_id_map or {}
      expected_single_punch_keys = set(expected_single_punch_keys or [])
      token = get_auth_token()

      # Reports most likely to expose BioTime's schedule-aware worked duration.
      report_names = (
          "totalTimeCardReportV2",
          "dailyActivityReport",
          "timeCardReport",
          "firstInLastOutReport",
          "monthlyWorkHoursReport",
      )
      query_variants = (
          {
              "start_date": start_text,
              "end_date": end_text,
              "page_size": 5000,
          },
          {
              "start_time": f"{start_text} 00:00:00",
              "end_time": f"{end_text} 23:59:59",
              "page_size": 5000,
          },
          {
              "from_date": start_text,
              "to_date": end_text,
              "page_size": 5000,
          },
      )

      # /att report endpoints work with HTTP Basic auth across BioTime editions.
      auth_attempts = [
          {
              "auth": (EMAIL, PASSWORD),
              "headers": {"Accept": "application/json"},
          }
      ]
      if token:
        auth_attempts.append(
            {
                "headers": {
                    "Authorization": f"Token {token}",
                    "Accept": "application/json",
                }
            }
        )

      best_records = {}
      best_scores = {}
      resolved_expected = set()

      def merge_payload(payload):
        nonlocal resolved_expected
        found = 0
        for report_row in extract_report_rows(payload):
          normalized_record = normalize_calculated_report_row(
              report_row,
              start_date,
              end_date,
              employee_internal_id_map,
          )
          if not normalized_record:
            continue

          key = (
              normalized_record["Attendance Date"],
              normalized_record["Employee ID"],
          )
          expected = key in expected_single_punch_keys

          # If single-punch keys are known from raw transactions, report rows for
          # those exact employee/date pairs are given the strongest priority.
          score = report_record_score(normalized_record, expected)
          if score > best_scores.get(key, -1):
            best_scores[key] = score
            best_records[key] = normalized_record

          if expected and normalized_record.get("Calculated WT"):
            resolved_expected.add(key)
          found += 1
        return found

      session = requests.Session()

      for report_name in report_names:
        endpoint = f"{BASE_URL}/att/api/{report_name}/"
        endpoint_missing = False

        for query in query_variants:
          if endpoint_missing:
            break

          query_returned_rows = False
          for authorization in auth_attempts:
            payloads, status_code = fetch_report_pages(
                session,
                endpoint,
                query,
                authorization,
            )
            if status_code == 404:
              endpoint_missing = True
              break
            if not payloads:
              continue

            rows_found = 0
            for payload in payloads:
              rows_found += merge_payload(payload)
            if rows_found:
              query_returned_rows = True
              break

          # Once a report responds with usable rows, do not repeat the same
          # endpoint with alternate date-parameter spellings.
          if query_returned_rows:
            break

        if expected_single_punch_keys and resolved_expected >= expected_single_punch_keys:
          break

      result = {}
      for (attendance_date, employee_id), record in best_records.items():
        result.setdefault(attendance_date, {})[employee_id] = record
      return result

    def fetch_paginated_api(session, path, params, headers, timeout=20):
      """Fetch a normal BioTime list endpoint, following pagination."""
      rows = []
      next_url = f"{BASE_URL}{path}"
      next_params = dict(params or {})
      seen_urls = set()

      for page_number in range(1, 50):
        request_key = (next_url, tuple(sorted((next_params or {}).items())))
        if request_key in seen_urls:
          break
        seen_urls.add(request_key)

        try:
          response = session.get(
              next_url,
              params=next_params,
              headers=headers,
              timeout=timeout,
          )
        except requests.RequestException:
          break

        if response.status_code != 200:
          break
        try:
          payload = response.json()
        except ValueError:
          break

        if isinstance(payload, list):
          rows.extend(item for item in payload if isinstance(item, dict))
          break
        if not isinstance(payload, dict):
          break

        page_rows = payload.get("data")
        if not isinstance(page_rows, list):
          page_rows = payload.get("results")
        if not isinstance(page_rows, list):
          page_rows = []
        rows.extend(item for item in page_rows if isinstance(item, dict))

        next_link = payload.get("next")
        if next_link:
          next_link = str(next_link)
          if next_link.startswith("http://") or next_link.startswith("https://"):
            next_url = next_link
          elif next_link.startswith("/"):
            next_url = BASE_URL + next_link
          else:
            next_url = BASE_URL + "/" + next_link.lstrip("/")
          next_params = None
          continue

        count = payload.get("count")
        try:
          count = int(count)
        except (TypeError, ValueError):
          count = len(rows)

        if len(rows) >= count or not page_rows:
          break

        # Some BioTime builds omit `next`; fall back to explicit page numbers.
        next_url = f"{BASE_URL}{path}"
        next_params = dict(params or {})
        next_params["page"] = page_number + 1

      return rows

    @strlit.cache_data(ttl=300, show_spinner=False)
    def load_monthly_attendance_bulk(start_date, end_date):
      """Load the whole attendance range with one request per data source.

      This replaces the old one-date-at-a-time loop, which repeatedly downloaded
      employees, devices, leave and transactions for every Excel date.
      """
      token = get_auth_token()
      if not token:
        raise RuntimeError("تعذر المصادقة مع BioTime")

      headers = {
          "Authorization": f"Token {token}",
          "Content-Type": "application/json",
      }
      session = requests.Session()

      devices = []
      for endpoint in ("/iclock/api/terminals/", "/iclock/api/devices/"):
        devices = fetch_paginated_api(
            session,
            endpoint,
            {"page_size": 1000},
            headers,
            timeout=15,
        )
        if devices:
          break

      terminal_map = {
          str(device.get("sn", "")): (
              device.get("alias")
              or device.get("terminal_name")
              or str(device.get("sn", ""))
          )
          for device in devices
          if device.get("sn")
      }

      all_employees = fetch_paginated_api(
          session,
          "/personnel/api/employees/",
          {"page_size": 1000},
          headers,
          timeout=20,
      )

      active_employees = {}
      employee_internal_id_map = {}

      for emp in all_employees:
        raw_code = str(emp.get("emp_code", "")).strip()
        cleaned_code = normalize_id(raw_code)

        is_active = (
            str(emp.get("is_active", True)).lower() in ("true", "1", "yes")
        )
        emp_status = str(emp.get("status", "0")).upper()
        enable_att = (
            str(emp.get("enable_attendance", True)).lower()
            in ("true", "1", "yes")
        )
        if not is_active or emp_status in ("1", "2", "D") or not enable_att:
          continue
        if not cleaned_code or cleaned_code in EXCLUDED_MANAGEMENT_CODES:
          continue

        first_name = str(emp.get("first_name", "") or "").strip()
        last_name = str(emp.get("last_name", "") or "").strip()
        if first_name.lower() == "none":
          first_name = ""
        if last_name.lower() == "none":
          last_name = ""
        full_name = f"{first_name} {last_name}".strip()

        dept_data = emp.get("department", {})
        dept_name = (
            dept_data.get("dept_name")
            if isinstance(dept_data, dict)
            else str(emp.get("department", "") or "")
        )
        if not dept_name or str(dept_name).lower() == "none":
          dept_name = "غير محدد"

        active_employees[cleaned_code] = {
            "name": clean_txt(full_name or f"موظف {cleaned_code}"),
            "dept": clean_txt(dept_name),
        }

        internal_id = normalize_id(emp.get("id"))
        if internal_id:
          employee_internal_id_map[internal_id] = cleaned_code

      employee_catalog = {
          employee_id: employee_data["name"]
          for employee_id, employee_data in active_employees.items()
      }

      leave_records = fetch_paginated_api(
          session,
          "/att/api/leave/",
          {"page_size": 5000},
          headers,
          timeout=15,
      )
      if not leave_records:
        leave_records = fetch_paginated_api(
            session,
            "/iclock/api/leave/",
            {"page_size": 5000},
            headers,
            timeout=15,
        )

      leave_by_date = {}
      for leave in leave_records:
        raw_code = (
            leave.get("emp_code")
            or leave.get("employee_code")
            or leave.get("employee_emp_code")
            or ""
        )
        cleaned_code = normalize_id(raw_code)
        if not cleaned_code:
          internal_employee = normalize_id(
              leave.get("employee_id") or leave.get("employee") or leave.get("emp")
          )
          cleaned_code = employee_internal_id_map.get(internal_employee, "")
        if cleaned_code not in active_employees:
          continue

        start_value = (
            leave.get("start_time")
            or leave.get("start_date")
            or leave.get("start_datetime")
        )
        end_value = (
            leave.get("end_time")
            or leave.get("end_date")
            or leave.get("end_datetime")
        )
        if not start_value or not end_value:
          continue

        try:
          leave_start = datetime.strptime(str(start_value)[:10], "%Y-%m-%d").date()
          leave_end = datetime.strptime(str(end_value)[:10], "%Y-%m-%d").date()
        except ValueError:
          continue

        leave_name = "إجازة"
        leave_type = leave.get("leave_type")
        if isinstance(leave_type, dict):
          leave_name = leave_type.get("leave_name") or leave_name
        elif leave_type not in (None, ""):
          leave_name = str(leave_type)
        elif leave.get("leave_name"):
          leave_name = str(leave.get("leave_name"))

        cursor = max(start_date, leave_start)
        last_date = min(end_date, leave_end)
        while cursor <= last_date:
          leave_by_date.setdefault(cursor, {})[cleaned_code] = clean_txt(leave_name)
          cursor += timedelta(days=1)

      raw_logs = fetch_paginated_api(
          session,
          "/iclock/api/transactions/",
          {
              "start_time": start_date.strftime("%Y-%m-%d") + " 00:00:00",
              "end_time": (end_date + timedelta(days=1)).strftime("%Y-%m-%d")
              + " 05:00:00",
              "page_size": 5000,
          },
          headers,
          timeout=30,
      )

      punches_by_emp_date = {}
      for log in raw_logs:
        cleaned_code = normalize_id(log.get("emp_code"))
        if not cleaned_code:
          internal_emp = normalize_id(log.get("emp"))
          cleaned_code = employee_internal_id_map.get(internal_emp, "")
        if cleaned_code not in active_employees:
          continue

        punch_value = log.get("punch_time")
        if not punch_value:
          continue
        try:
          punch_time = datetime.strptime(
              str(punch_value)[:19],
              "%Y-%m-%d %H:%M:%S",
          )
        except ValueError:
          continue

        device_sn = str(log.get("terminal_sn", "") or "")
        device_name = (
            log.get("terminal_alias")
            or log.get("terminal_name")
            or terminal_map.get(device_sn, device_sn or "جهاز رئيسي")
        )
        punches_by_emp_date.setdefault(
            (cleaned_code, punch_time.date()), []
        ).append((punch_time, device_name))

      # Remove duplicate punches (within 60 seconds) once for the whole month.
      for key, punches in list(punches_by_emp_date.items()):
        punches = sorted(punches, key=lambda item: item[0])
        filtered = []
        for punch_time, device_name in punches:
          if (
              not filtered
              or abs((punch_time - filtered[-1][0]).total_seconds()) > 60
          ):
            filtered.append((punch_time, device_name))
        punches_by_emp_date[key] = filtered

      all_attendance = {}
      today = datetime.now(SYRIA_TZ).date()
      cursor = start_date

      while cursor <= end_date:
        date_map = {}
        is_date_today = cursor == today
        leave_map = leave_by_date.get(cursor, {})

        for code, emp_data in active_employees.items():
          name = emp_data["name"]
          dept = emp_data["dept"]

          day_punches = [
              (punch_time, device_name)
              for punch_time, device_name in punches_by_emp_date.get(
                  (code, cursor), []
              )
              if punch_time.hour >= 5
          ]
          next_morning = [
              (punch_time, device_name)
              for punch_time, device_name in punches_by_emp_date.get(
                  (code, cursor + timedelta(days=1)), []
              )
              if punch_time.hour < 5
          ]

          if not day_punches:
            if code in leave_map:
              date_map[code] = {
                  "Employee ID": code,
                  "First Name": name,
                  "Department": dept,
                  "Date": cursor.strftime("%Y-%m-%d"),
                  "Clock In": "",
                  "Clock Out": "",
                  "Total WT": "",
                  "Status": f"Leave - {leave_map[code]}",
              }
            else:
              date_map[code] = {
                  "Employee ID": code,
                  "First Name": name,
                  "Department": dept,
                  "Date": cursor.strftime("%Y-%m-%d"),
                  "Clock In": "",
                  "Clock Out": "",
                  "Total WT": "",
                  "Status": "Absence(A)",
              }
            continue

          first_p, first_dev = day_punches[0]
          single_punch_only = len(day_punches) == 1 and not next_morning
          single_punch_is_out = (
              single_punch_only and first_p.hour >= SINGLE_PUNCH_OUT_HOUR
          )
          is_late = (
              not single_punch_is_out
              and (first_p.hour > 9 or (first_p.hour == 9 and first_p.minute > 15))
          )

          if single_punch_only:
            single_punch_actual_wt = calculate_single_punch_actual_wt(
                first_p,
                single_punch_is_out,
            )
            date_map[code] = {
                "Employee ID": code,
                "First Name": name,
                "Department": dept,
                "Date": cursor.strftime("%Y-%m-%d"),
                "Clock In": "" if single_punch_is_out else first_p.strftime("%H:%M"),
                "Clock Out": first_p.strftime("%H:%M") if single_punch_is_out else "",
                "Actual WT": single_punch_actual_wt,
                "Calculated WT": single_punch_actual_wt,
                "Total WT": single_punch_actual_wt,
                "Status": (
                    "Present(P) / Missing IN"
                    if single_punch_is_out
                    else (
                        "Late(LT) / Missing OUT"
                        if is_late
                        else "Present(P) / Missing OUT"
                    )
                ),
            }
            continue

          punch_count = (
              2 if (len(day_punches) % 2 != 0 and next_morning) else len(day_punches)
          )
          last_p = None
          if punch_count % 2 == 0:
            last_p = (
                next_morning[-1][0]
                if (len(day_punches) % 2 != 0 and next_morning)
                else day_punches[-1][0]
            )
          elif not is_date_today and len(day_punches) > 1:
            last_p = day_punches[-1][0]

          total_work = ""
          if last_p:
            diff = last_p - first_p
            total_seconds = max(0, int(diff.total_seconds()))
            hours, remainder = divmod(total_seconds, 3600)
            minutes = remainder // 60
            total_work = f"{hours:02d}:{minutes:02d}"

          date_map[code] = {
              "Employee ID": code,
              "First Name": name,
              "Department": dept,
              "Date": cursor.strftime("%Y-%m-%d"),
              "Clock In": first_p.strftime("%H:%M"),
              "Clock Out": last_p.strftime("%H:%M") if last_p else "",
              "Total WT": total_work,
              "Status": "Late(LT)" if is_late else "Present(P)",
          }

        all_attendance[cursor] = date_map
        cursor += timedelta(days=1)

      return all_attendance, employee_catalog, employee_internal_id_map

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

        # Attendance values feed formulas on this sheet and the payroll sheet.
        # The original package contains cached formula results (often zero).
        # Force Excel to recalculate ALL formulas on open so total hours, total
        # attendance days, overtime/payroll links and subtotals refresh correctly.
        workbook_xml_text = zin.read("xl/workbook.xml").decode("utf-8")
        calc_pattern = re.compile(r"<calcPr\b([^>]*)/>")
        calc_match = calc_pattern.search(workbook_xml_text)
        if calc_match:
          calc_attrs = calc_match.group(1)
          calc_attrs = re.sub(
              r'\s+(?:calcMode|calcOnSave|fullCalcOnLoad|forceFullCalc)="[^"]*"',
              "",
              calc_attrs,
          )
          calc_replacement = (
              f'<calcPr{calc_attrs} calcMode="auto" calcOnSave="1" '
              'fullCalcOnLoad="1" forceFullCalc="1"/>'
          )
          workbook_xml_text = (
              workbook_xml_text[:calc_match.start()]
              + calc_replacement
              + workbook_xml_text[calc_match.end():]
          )
        else:
          workbook_xml_text = workbook_xml_text.replace(
              "</workbook>",
              '<calcPr calcMode="auto" calcOnSave="1" fullCalcOnLoad="1" '
              'forceFullCalc="1"/></workbook>',
          )
        patched_workbook_xml = workbook_xml_text.encode("utf-8")

        with zipfile.ZipFile(output_buffer, "w") as zout:
          for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == sheet_xml_path:
              data = patched_sheet_xml
            elif item.filename == "xl/workbook.xml":
              data = patched_workbook_xml
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

        # FAST MONTHLY LOAD:
        # Employees, devices, leave and transactions are fetched once for the
        # entire Excel date range. Single-punch Actual WT is calculated locally
        # with the same 09:00-19:00 timetable rule verified against BioTime's
        # Basic Report, so there is no slow report-API probing and no second file.
        sorted_dates = sorted(set(date_columns.values()))
        dates_list = [
            (column, attendance_date)
            for column, attendance_date in date_columns.items()
            if attendance_date <= now_syria.date()
        ]

        if not dates_list:
          strlit.error("لا توجد تواريخ حالية قابلة للمعالجة في ملف الدوام.")
          raise RuntimeError("No current attendance dates to process")

        range_start = min(attendance_date for _column, attendance_date in dates_list)
        range_end = max(attendance_date for _column, attendance_date in dates_list)

        progress = strlit.progress(
            0,
            text="جاري تحميل بيانات الدوام الشهرية...",
        )

        all_attendance, employee_catalog, _employee_internal_id_map = (
            load_monthly_attendance_bulk(range_start, range_end)
        )
        progress.progress(
            0.95,
            text="تم تحميل البصمات. جاري تجهيز الملف...",
        )

        recovered_single_punch = 0
        for attendance_date, date_map in all_attendance.items():
          for employee_id, attendance in date_map.items():
            status = str(attendance.get("Status", "") or "")
            if "Leave" in status or "Absence" in status:
              continue
            clock_in = str(attendance.get("Clock In", "") or "").strip()
            clock_out = str(attendance.get("Clock Out", "") or "").strip()
            total_work = str(
                attendance.get("Calculated WT", "")
                or attendance.get("Actual WT", "")
                or attendance.get("Total WT", "")
                or ""
            ).strip()
            if bool(clock_in) != bool(clock_out) and total_work:
              recovered_single_punch += 1

        progress.progress(
            1.0,
            text="تم تجهيز بيانات الدوام.",
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
                total_work = str(
                    attendance.get("Calculated WT", "")
                    or attendance.get("Actual WT", "")
                    or attendance.get("Total WT", "")
                    or ""
                ).strip()

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
              "single_punch": recovered_single_punch,
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
