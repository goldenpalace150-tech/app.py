import streamlit as strlit
import requests
import unicodedata
import pandas as pd
import io
import base64
from datetime import datetime, timedelta
import zoneinfo
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# 0. PAGE / TEXT CONFIGURATION
# ============================================================

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


# ============================================================
# 1. CSS
# ============================================================

strlit.markdown(
    """
    <style>

    header[data-testid="stHeader"] {
        display: none !important;
    }

    footer {
        display: none !important;
    }

    .stApp {
        direction: rtl;
        background-color: #f4f7f9;
        font-family: system-ui, -apple-system, sans-serif;
    }

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
        0% {
            transform: rotate(0deg);
        }

        100% {
            transform: rotate(360deg);
        }
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
        0% {
            transform: scale(0.95);
            box-shadow: 0 0 0 0 rgba(34, 197, 94, 0.7);
        }

        70% {
            transform: scale(1);
            box-shadow: 0 0 0 6px rgba(34, 197, 94, 0);
        }

        100% {
            transform: scale(0.95);
            box-shadow: 0 0 0 0 rgba(34, 197, 94, 0);
        }
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
        padding: 10px;
        border-bottom: 1px solid #f1f5f9;
        text-align: center;
        font-weight: 500;
        color: #1e293b;
    }

    .badge-present {
        background-color: #dcfce7;
        color: #166534;
        padding: 4px 8px;
        border-radius: 6px;
        font-size: 11px;
    }

    .badge-late {
        background-color: #fef3c7;
        color: #9a3412;
        padding: 4px 8px;
        border-radius: 6px;
        font-size: 11px;
    }

    .badge-leave {
        background-color: #e0f2fe;
        color: #0369a1;
        padding: 4px 8px;
        border-radius: 6px;
        font-size: 11px;
    }

    .badge-absent {
        background-color: #fee2e2;
        color: #991b1b;
        padding: 4px 8px;
        border-radius: 6px;
        font-size: 11px;
    }

    .api-error-box {
        direction: rtl;
        text-align: right;
        padding: 16px;
        border-radius: 12px;
        border: 1px solid #fecaca;
        background: #fef2f2;
        color: #991b1b;
        margin-top: 15px;
        line-height: 1.8;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 2. BASIC CONFIGURATION
# ============================================================

EXCLUDED_MANAGEMENT_CODES = ("40",)

SYRIA_TZ = zoneinfo.ZoneInfo("Asia/Damascus")


# ============================================================
# 3. BIOTIME SECRETS
# ============================================================
#
# REQUIRED:
#
# [biotime]
# base_url = "http://YOUR-BIOTIME-SERVER:8090"
# username = "YOUR_BIOTIME_USERNAME"
# password = "YOUR_BIOTIME_PASSWORD"
#
# OPTIONAL:
#
# token_url = "http://YOUR-BIOTIME-SERVER:8090/api-token-auth/"
#
# The code also accepts the old "email" key as a fallback for
# compatibility, but BioTime authentication itself uses it as
# the username.
# ============================================================

try:
    BASE_URL = strlit.secrets["biotime"]["base_url"].rstrip("/")

    BIOTIME_USERNAME = strlit.secrets["biotime"].get(
        "username",
        strlit.secrets["biotime"].get("email", "")
    )

    BIOTIME_PASSWORD = strlit.secrets["biotime"]["password"]

    TOKEN_URL = strlit.secrets["biotime"].get(
        "token_url",
        f"{BASE_URL}/api-token-auth/"
    )

except Exception as config_error:
    strlit.error(
        "BioTime configuration is missing from Streamlit Secrets. "
        f"Details: {config_error}"
    )
    strlit.stop()


if not BIOTIME_USERNAME:
    strlit.error(
        "BioTime username is missing. Add 'username' to the [biotime] "
        "section in Streamlit Secrets."
    )
    strlit.stop()


# ============================================================
# 4. SESSION STATE
# ============================================================

if "debug_logs" not in strlit.session_state:
    strlit.session_state["debug_logs"] = []

if "selected_view" not in strlit.session_state:
    strlit.session_state["selected_view"] = "present"

if "last_selected_date" not in strlit.session_state:
    strlit.session_state["last_selected_date"] = None


# ============================================================
# 5. HELPER FUNCTIONS
# ============================================================

def clean_txt(raw_text):
    if raw_text is None:
        return ""

    try:
        text = str(raw_text)
        text = unicodedata.normalize("NFKC", text)
        text = (
            text
            .replace("\u2066", "")
            .replace("\u2069", "")
            .strip()
        )
        return text
    except Exception:
        return str(raw_text).strip()


def clean_employee_code(raw_code):
    """
    Normalize employee codes such as:
    0012 -> 12
    12   -> 12
    ABC  -> ABC
    """
    if raw_code is None:
        return ""

    value = str(raw_code).strip()

    if value.isdigit():
        try:
            return str(int(value))
        except Exception:
            return value

    return value


def safe_json(response):
    try:
        return response.json()
    except Exception:
        return {}


def extract_token(data):
    """
    BioTime versions may expose token information differently.
    Check the common fields without exposing the token.
    """

    if not isinstance(data, dict):
        return None

    possible_fields = [
        "token",
        "key",
        "access",
    ]

    for field in possible_fields:
        value = data.get(field)

        if value:
            return str(value).strip()

    return None


# ============================================================
# 6. BIOTIME AUTHENTICATION
# ============================================================

@strlit.cache_data(ttl=300)
def get_auth_token():

    payload = {
        "username": BIOTIME_USERNAME,
        "password": BIOTIME_PASSWORD,
    }

    # First try the configured endpoint.
    endpoints_to_try = []

    if TOKEN_URL:
        endpoints_to_try.append(TOKEN_URL.rstrip("/") + "/")

    # Always keep the documented standard endpoint as fallback.
    standard_endpoint = f"{BASE_URL}/api-token-auth/"

    if standard_endpoint not in endpoints_to_try:
        endpoints_to_try.append(standard_endpoint)

    last_error = None

    for endpoint in endpoints_to_try:

        try:

            response = requests.post(
                endpoint,
                json=payload,
                headers={
                    "Content-Type": "application/json",
                    "Accept": "application/json",
                },
                timeout=15,
            )

        except requests.exceptions.Timeout:
            last_error = (
                f"Connection timeout while contacting BioTime at "
                f"{endpoint}"
            )
            continue

        except requests.exceptions.ConnectionError as exc:
            last_error = (
                f"Cannot connect to BioTime at {endpoint}. "
                f"Network error: {str(exc)}"
            )
            continue

        except requests.exceptions.RequestException as exc:
            last_error = (
                f"HTTP request error while contacting BioTime: "
                f"{str(exc)}"
            )
            continue

        data = safe_json(response)

        # ----------------------------------------------------
        # SUCCESS
        # ----------------------------------------------------

        if response.status_code in (200, 201):

            token = extract_token(data)

            if token:
                return token

            last_error = (
                f"BioTime accepted the login request with HTTP "
                f"{response.status_code}, but no token was found "
                f"in the response."
            )

            continue

        # ----------------------------------------------------
        # AUTHENTICATION ERROR
        # ----------------------------------------------------

        if response.status_code in (400, 401, 403):

            detail = ""

            if isinstance(data, dict):
                detail = (
                    data.get("detail")
                    or data.get("msg")
                    or data.get("message")
                    or data.get("error")
                    or ""
                )

            if not detail:
                detail = response.text[:500]

            last_error = (
                f"BioTime authentication failed. "
                f"HTTP {response.status_code}. "
                f"Server response: {detail}"
            )

            # Don't keep trying another endpoint for invalid credentials.
            break

        # ----------------------------------------------------
        # NOT FOUND
        # ----------------------------------------------------

        if response.status_code == 404:

            last_error = (
                f"BioTime token endpoint was not found: {endpoint}. "
                f"HTTP 404."
            )

            continue

        # ----------------------------------------------------
        # OTHER SERVER ERROR
        # ----------------------------------------------------

        last_error = (
            f"BioTime returned HTTP {response.status_code}. "
            f"Response: {response.text[:500]}"
        )

    raise Exception(
        last_error
        or "Unable to obtain a BioTime authentication token."
    )


# ============================================================
# 7. GENERIC BIOTIME GET
# ============================================================

def biotime_get(endpoint, token, timeout=15, params=None):

    url = f"{BASE_URL}{endpoint}"

    response = requests.get(
        url,
        headers={
            "Authorization": f"Token {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        params=params,
        timeout=timeout,
    )

    if response.status_code == 401:

        raise Exception(
            "BioTime rejected the authentication token "
            "(HTTP 401). The token may have expired or the "
            "BioTime user may not have API access."
        )

    if response.status_code == 403:

        raise Exception(
            "BioTime denied access to the requested API "
            f"endpoint (HTTP 403): {endpoint}"
        )

    if response.status_code != 200:

        raise Exception(
            f"BioTime API returned HTTP {response.status_code} "
            f"for {endpoint}. "
            f"Response: {response.text[:500]}"
        )

    return safe_json(response)


# ============================================================
# 8. LOAD ATTENDANCE DATA
# ============================================================

def load_attendance_data_from_api(
    selected_date_str,
    selected_date_obj,
    is_today
):

    # --------------------------------------------------------
    # AUTHENTICATION
    # --------------------------------------------------------

    token = get_auth_token()

    if not token:
        raise Exception(
            "BioTime authentication token was not returned."
        )

    # --------------------------------------------------------
    # 1. DEVICES
    # --------------------------------------------------------

    devices = []

    device_errors = []

    for endpoint in [
        "/iclock/api/terminals/",
        "/iclock/api/devices/",
    ]:

        try:

            data = biotime_get(
                endpoint,
                token,
                timeout=10,
                params={"page_size": 1000},
            )

            if isinstance(data, dict):

                devices = data.get(
                    "data",
                    data.get("results", [])
                )

            elif isinstance(data, list):

                devices = data

            if isinstance(devices, list):
                break

        except Exception as exc:
            device_errors.append(str(exc))

    if not isinstance(devices, list):
        devices = []

    terminal_map = {}

    for device in devices:

        if not isinstance(device, dict):
            continue

        sn = str(device.get("sn", "")).strip()

        if not sn:
            continue

        alias = (
            device.get("alias")
            or device.get("terminal_name")
            or sn
        )

        terminal_map[sn] = clean_txt(alias)

    # --------------------------------------------------------
    # 2. EMPLOYEES
    # --------------------------------------------------------

    all_employees = []

    try:

        employee_data = biotime_get(
            "/personnel/api/employees/",
            token,
            timeout=15,
            params={"page_size": 1000},
        )

        if isinstance(employee_data, dict):

            all_employees = employee_data.get(
                "data",
                employee_data.get("results", [])
            )

        elif isinstance(employee_data, list):

            all_employees = employee_data

    except Exception as exc:

        raise Exception(
            f"Unable to load employees from BioTime: {exc}"
        )

    if not isinstance(all_employees, list):
        all_employees = []

    active_employees = {}

    for emp in all_employees:

        if not isinstance(emp, dict):
            continue

        raw_code = emp.get("emp_code", "")
        cleaned_code = clean_employee_code(raw_code)

        if not cleaned_code:
            continue

        is_active = str(
            emp.get("is_active", True)
        ).lower() in ("true", "1", "yes")

        emp_status = str(
            emp.get("status", "0")
        ).upper()

        enable_att = str(
            emp.get("enable_attendance", True)
        ).lower() in ("true", "1", "yes")

        if (
            not is_active
            or emp_status in ("1", "2", "D")
            or not enable_att
        ):
            continue

        if cleaned_code in EXCLUDED_MANAGEMENT_CODES:
            continue

        first_name = str(
            emp.get("first_name", "")
        ).strip()

        last_name = str(
            emp.get("last_name", "")
        ).strip()

        if first_name.lower() == "none":
            first_name = ""

        if last_name.lower() == "none":
            last_name = ""

        full_name = (
            f"{first_name} {last_name}"
        ).strip()

        # Department
        department_data = emp.get("department", {})

        if isinstance(department_data, dict):

            department_name = (
                department_data.get("dept_name")
                or department_data.get("name")
                or ""
            )

        else:

            department_name = str(
                department_data or ""
            )

        if (
            not department_name
            or department_name.lower() == "none"
        ):
            department_name = "غير محدد"

        active_employees[cleaned_code] = {
            "name": clean_txt(
                full_name
                if full_name
                else f"موظف {cleaned_code}"
            ),
            "dept": clean_txt(department_name),
        }

    # --------------------------------------------------------
    # 3. LEAVES
    # --------------------------------------------------------

    leave_records = []

    leave_endpoints = [
        "/att/api/leave/",
        "/iclock/api/leave/",
    ]

    for endpoint in leave_endpoints:

        try:

            leave_data = biotime_get(
                endpoint,
                token,
                timeout=10,
                params={"page_size": 1000},
            )

            if isinstance(leave_data, dict):

                leave_records = leave_data.get(
                    "data",
                    leave_data.get("results", [])
                )

            elif isinstance(leave_data, list):

                leave_records = leave_data

            if isinstance(leave_records, list):
                break

        except Exception:
            continue

    if not isinstance(leave_records, list):
        leave_records = []

    on_leave_employees = {}

    for leave in leave_records:

        if not isinstance(leave, dict):
            continue

        raw_code = (
            leave.get("emp_code")
            or leave.get("employee_code")
            or ""
        )

        cleaned_code = clean_employee_code(raw_code)

        start_time = (
            leave.get("start_time")
            or leave.get("start_date")
            or leave.get("start_datetime")
        )

        end_time = (
            leave.get("end_time")
            or leave.get("end_date")
            or leave.get("end_datetime")
        )

        if not cleaned_code or not start_time or not end_time:
            continue

        leave_name = "إجازة"

        if "leave_type" in leave:

            leave_type = leave["leave_type"]

            if isinstance(leave_type, dict):

                leave_name = (
                    leave_type.get("leave_name")
                    or leave_type.get("name")
                    or "إجازة"
                )

            else:

                leave_name = str(leave_type)

        elif "leave_name" in leave:

            leave_name = str(
                leave.get("leave_name")
                or "إجازة"
            )

        try:

            start_string = str(start_time)[:10]
            end_string = str(end_time)[:10]

            start_date = datetime.strptime(
                start_string,
                "%Y-%m-%d"
            ).date()

            end_date = datetime.strptime(
                end_string,
                "%Y-%m-%d"
            ).date()

            if start_date <= selected_date_obj <= end_date:

                on_leave_employees[
                    cleaned_code
                ] = clean_txt(leave_name)

        except Exception:
            continue

    # --------------------------------------------------------
    # 4. TRANSACTION LOGS
    # --------------------------------------------------------

    start_datetime = (
        selected_date_obj.strftime("%Y-%m-%d")
        + " 00:00:00"
    )

    end_datetime = (
        (
            selected_date_obj
            + timedelta(days=1)
        ).strftime("%Y-%m-%d")
        + " 05:00:00"
    )

    raw_logs = []

    try:

        transaction_data = biotime_get(
            "/iclock/api/transactions/",
            token,
            timeout=20,
            params={
                "start_time": start_datetime,
                "end_time": end_datetime,
                "page_size": 5000,
            },
        )

        if isinstance(transaction_data, dict):

            raw_logs = transaction_data.get(
                "data",
                transaction_data.get("results", [])
            )

        elif isinstance(transaction_data, list):

            raw_logs = transaction_data

    except Exception as exc:

        raise Exception(
            f"Unable to load attendance transactions "
            f"from BioTime: {exc}"
        )

    if not isinstance(raw_logs, list):
        raw_logs = []

    # --------------------------------------------------------
    # ORGANIZE EMPLOYEE PUNCHES
    # --------------------------------------------------------

    emp_punches = {}

    for log in raw_logs:

        if not isinstance(log, dict):
            continue

        raw_code = log.get(
            "emp_code",
            ""
        )

        cleaned_code = clean_employee_code(raw_code)

        if cleaned_code not in active_employees:
            continue

        punch_time_raw = log.get("punch_time")

        if not punch_time_raw:
            continue

        try:

            punch_string = str(
                punch_time_raw
            )[:19]

            punch_time = datetime.strptime(
                punch_string,
                "%Y-%m-%d %H:%M:%S"
            )

            device_sn = str(
                log.get(
                    "terminal_sn",
                    ""
                )
            )

            device_name = (
                log.get("terminal_alias")
                or log.get("terminal_name")
                or terminal_map.get(
                    device_sn,
                    device_sn or "جهاز رئيسي"
                )
            )

            emp_punches.setdefault(
                cleaned_code,
                []
            ).append(
                (
                    punch_time,
                    clean_txt(device_name)
                )
            )

        except Exception:
            continue

    # --------------------------------------------------------
    # OUTPUT COLLECTIONS
    # --------------------------------------------------------

    present_staff = []
    late_staff = []
    absent_staff = []
    checkout_staff = []
    leave_staff = []
    excel_rows = []

    # --------------------------------------------------------
    # PROCESS EACH ACTIVE EMPLOYEE
    # --------------------------------------------------------

    for code, employee_data in active_employees.items():

        name = employee_data["name"]
        department = employee_data["dept"]

        punches = sorted(
            emp_punches.get(code, []),
            key=lambda item: item[0]
        )

        # Remove duplicate punches within 60 seconds.
        filtered_punches = []

        for punch_time, device_name in punches:

            if (
                not filtered_punches
                or abs(
                    (
                        punch_time
                        - filtered_punches[-1][0]
                    ).total_seconds()
                ) > 60
            ):

                filtered_punches.append(
                    (
                        punch_time,
                        device_name
                    )
                )

        # Only punches after 05:00 on the selected date.
        day_punches = [
            (punch, device)
            for punch, device in filtered_punches
            if (
                punch.date() == selected_date_obj
                and punch.hour >= 5
            )
        ]

        # ----------------------------------------------------
        # NO PUNCH
        # ----------------------------------------------------

        if not day_punches:

            if code in on_leave_employees:

                leave_reason = on_leave_employees[code]

                leave_staff.append(
                    (
                        code,
                        name,
                        department,
                        leave_reason
                    )
                )

                excel_rows.append({
                    "Employee ID": code,
                    "First Name": name,
                    "Department": department,
                    "Date": selected_date_str,
                    "Clock In": "",
                    "Clock Out": "",
                    "Total WT": "",
                    "Status": f"Leave - {leave_reason}",
                })

            else:

                absent_staff.append(
                    (
                        code,
                        name,
                        department
                    )
                )

                excel_rows.append({
                    "Employee ID": code,
                    "First Name": name,
                    "Department": department,
                    "Date": selected_date_str,
                    "Clock In": "",
                    "Clock Out": "",
                    "Total WT": "",
                    "Status": "Absence(A)",
                })

            continue

        # ----------------------------------------------------
        # FIRST PUNCH
        # ----------------------------------------------------

        first_punch, first_device = day_punches[0]

        is_late = (
            first_punch.hour > 9
            or (
                first_punch.hour == 9
                and first_punch.minute > 15
            )
        )

        # ----------------------------------------------------
        # NEXT MORNING PUNCHES
        # ----------------------------------------------------

        next_morning = [
            (punch, device)
            for punch, device in filtered_punches
            if (
                punch.date()
                == selected_date_obj + timedelta(days=1)
                and punch.hour < 5
            )
        ]

        # ----------------------------------------------------
        # DETERMINE PUNCH COUNT
        # ----------------------------------------------------

        punch_count = len(day_punches)

        if (
            len(day_punches) % 2 != 0
            and next_morning
        ):
            punch_count = 2

        # ----------------------------------------------------
        # DETERMINE LAST PUNCH
        # ----------------------------------------------------

        last_punch = None
        last_device = first_device

        if punch_count % 2 == 0:

            if (
                len(day_punches) % 2 != 0
                and next_morning
            ):

                last_punch, last_device = (
                    next_morning[-1]
                )

            else:

                last_punch, last_device = (
                    day_punches[-1]
                )

        elif not is_today and len(day_punches) > 1:

            last_punch, last_device = (
                day_punches[-1]
            )

        # ----------------------------------------------------
        # TOTAL WORK TIME
        # ----------------------------------------------------

        total_work_time = ""

        if last_punch and first_punch:

            difference = (
                last_punch - first_punch
            )

            total_seconds = int(
                difference.total_seconds()
            )

            if total_seconds >= 0:

                hours, remainder = divmod(
                    total_seconds,
                    3600
                )

                minutes = remainder // 60

                total_work_time = (
                    f"{hours:02d}:{minutes:02d}"
                )

        # ----------------------------------------------------
        # STATUS
        # ----------------------------------------------------

        status_string = (
            "Late(LT)"
            if is_late
            else "Present(P)"
        )

        if is_late:

            late_staff.append(
                (
                    code,
                    name,
                    department,
                    first_punch.strftime("%I:%M %p"),
                    first_device,
                )
            )

        # ====================================================
        # TODAY
        # ====================================================

        if is_today:

            # Odd number of punches = currently inside.
            if punch_count % 2 != 0:

                present_staff.append(
                    (
                        code,
                        name,
                        department,
                        first_punch.strftime("%I:%M %p"),
                        first_device,
                    )
                )

                excel_rows.append({
                    "Employee ID": code,
                    "First Name": name,
                    "Department": department,
                    "Date": selected_date_str,
                    "Clock In": first_punch.strftime("%H:%M"),
                    "Clock Out": "",
                    "Total WT": "",
                    "Status": status_string,
                })

            # Even number of punches = checked out.
            else:

                if (
                    len(day_punches) % 2 != 0
                    and next_morning
                ):

                    last_real_punch, last_real_device = (
                        next_morning[-1]
                    )

                else:

                    last_real_punch, last_real_device = (
                        day_punches[-1]
                    )

                checkout_staff.append(
                    (
                        code,
                        name,
                        department,
                        last_real_punch.strftime("%I:%M %p"),
                        last_real_device,
                    )
                )

                excel_rows.append({
                    "Employee ID": code,
                    "First Name": name,
                    "Department": department,
                    "Date": selected_date_str,
                    "Clock In": first_punch.strftime("%H:%M"),
                    "Clock Out": last_real_punch.strftime("%H:%M"),
                    "Total WT": total_work_time,
                    "Status": status_string,
                })

        # ====================================================
        # HISTORICAL DATE
        # ====================================================

        else:

            if last_punch:

                checkout_staff.append(
                    (
                        code,
                        name,
                        department,
                        last_punch.strftime("%I:%M %p"),
                        last_device,
                    )
                )

            else:

                present_staff.append(
                    (
                        code,
                        name,
                        department,
                        first_punch.strftime("%I:%M %p"),
                        first_device,
                    )
                )

            excel_rows.append({
                "Employee ID": code,
                "First Name": name,
                "Department": department,
                "Date": selected_date_str,
                "Clock In": first_punch.strftime("%H:%M"),
                "Clock Out": (
                    last_punch.strftime("%H:%M")
                    if last_punch
                    else ""
                ),
                "Total WT": total_work_time,
                "Status": status_string,
            })

    # --------------------------------------------------------
    # SORT
    # --------------------------------------------------------

    def sort_key(item):
        code = str(item[0])

        if code.isdigit():
            return (0, int(code))

        return (1, code.lower())

    absent_staff.sort(key=sort_key)
    present_staff.sort(key=sort_key)
    late_staff.sort(key=sort_key)
    leave_staff.sort(key=sort_key)
    checkout_staff.sort(key=sort_key)

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


# ============================================================
# 9. CURRENT TIME
# ============================================================

now_syria = datetime.now(SYRIA_TZ)

today_str = now_syria.strftime("%Y-%m-%d")


# ============================================================
# 10. HEADER IMAGE
# ============================================================

dish_img_tag = ""

try:

    with open(
        "image_632b3d.jpg",
        "rb"
    ) as image_file:

        encoded_string = base64.b64encode(
            image_file.read()
        ).decode()

        dish_img_tag = (
            '<img '
            f'src="data:image/jpeg;base64,{encoded_string}" '
            'class="animated-dish" />'
        )

except Exception:

    dish_img_tag = (
        '<div '
        'class="animated-dish" '
        'style="font-size: 24px;">'
        '📡'
        '</div>'
    )


# ============================================================
# 11. DATE + REFRESH
# ============================================================

column_date, column_refresh = strlit.columns(2)

with column_date:

    selected_date_obj_input = strlit.date_input(
        "",
        value=now_syria.date(),
        label_visibility="collapsed",
    )

    selected_date_str = (
        selected_date_obj_input.strftime("%Y-%m-%d")
    )

with column_refresh:

    if strlit.button(
        "🔄 تحديث البيانات",
        use_container_width=True
    ):

        strlit.cache_data.clear()
        strlit.rerun()


is_today = (
    selected_date_str == today_str
)


# ============================================================
# 12. RESET VIEW WHEN DATE CHANGES
# ============================================================

if (
    strlit.session_state["last_selected_date"]
    is None
):

    strlit.session_state[
        "last_selected_date"
    ] = selected_date_str


if (
    strlit.session_state["last_selected_date"]
    != selected_date_str
):

    strlit.session_state[
        "last_selected_date"
    ] = selected_date_str

    strlit.session_state[
        "selected_view"
    ] = (
        "present"
        if is_today
        else "all"
    )


# ============================================================
# 13. ONLINE / ARCHIVE STATUS
# ============================================================

if is_today:

    strlit.markdown(
        f"""
        <div class="status-badge">

            {dish_img_tag}

            <div class="status-indicator">

                <span class="blinking-dot"></span>

                <span class="online-text">
                    Online (مباشر)
                </span>

            </div>

        </div>
        """,
        unsafe_allow_html=True,
    )

else:

    strlit.markdown(
        f"""
        <div
            class="status-badge"
            style="
                border-color: #94a3b8;
                background: #e2e8f0;
            "
        >

            {dish_img_tag}

            <div class="status-indicator">

                <span
                    style="
                        font-weight: 800;
                        color: #475569;
                        font-size: 14px;
                    "
                >
                    أرشيف تاريخي ({selected_date_str})
                </span>

            </div>

        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# 14. LOAD DATA
# ============================================================

try:

    (
        active_employees,
        present_staff,
        late_staff,
        absent_staff,
        checkout_staff,
        leave_staff,
        devices,
        excel_rows,
    ) = load_attendance_data_from_api(
        selected_date_str,
        selected_date_obj_input,
        is_today,
    )

    # ========================================================
    # 15. EXCEL REPORT
    # ========================================================

    df_excel = pd.DataFrame(excel_rows)

    output = io.BytesIO()

    workbook = openpyxl.Workbook()

    worksheet = workbook.active

    worksheet.title = "Attendance Report"

    thin_border = Border(
        left=Side(
            style="thin",
            color="D3D3D3"
        ),
        right=Side(
            style="thin",
            color="D3D3D3"
        ),
        top=Side(
            style="thin",
            color="D3D3D3"
        ),
        bottom=Side(
            style="thin",
            color="D3D3D3"
        ),
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

    worksheet.append(headers)

    worksheet.row_dimensions[1].height = 24

    for column_index in range(
        1,
        len(headers) + 1
    ):

        cell = worksheet.cell(
            row=1,
            column=column_index
        )

        cell.font = Font(
            name="Calibri",
            size=11,
            bold=True,
            color="FFFFFF",
        )

        cell.fill = PatternFill(
            start_color="1F4E78",
            end_color="1F4E78",
            fill_type="solid",
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

        cell.border = thin_border

    for row_index, row_data in enumerate(
        excel_rows,
        2
    ):

        worksheet.row_dimensions[
            row_index
        ].height = 20

        worksheet.append([
            row_data["Employee ID"],
            row_data["First Name"],
            row_data["Department"],
            row_data["Date"],
            row_data["Clock In"],
            row_data["Clock Out"],
            row_data["Total WT"],
            row_data["Status"],
        ])

        status_value = str(
            row_data["Status"]
        )

        row_fill = None
        row_font_color = "000000"

        if (
            "Leave" in status_value
            or status_value == "L"
        ):

            row_fill = PatternFill(
                start_color="D9E1F2",
                end_color="D9E1F2",
                fill_type="solid",
            )

            row_font_color = "002060"

        elif (
            "Absence" in status_value
            or status_value == "A"
        ):

            row_fill = PatternFill(
                start_color="FFC7CE",
                end_color="FFC7CE",
                fill_type="solid",
            )

            row_font_color = "9C0006"

        for column_index in range(1, 9):

            cell = worksheet.cell(
                row=row_index,
                column=column_index
            )

            if row_fill:

                cell.fill = row_fill

                cell.font = Font(
                    name="Calibri",
                    size=11,
                    bold=True,
                    color=row_font_color,
                )

            else:

                cell.font = Font(
                    name="Calibri",
                    size=11
                )

            # Status column
            if (
                column_index == 8
                and not row_fill
            ):

                if (
                    "Late" in status_value
                    or "LT" in status_value
                ):

                    cell.font = Font(
                        name="Calibri",
                        size=11,
                        bold=True,
                        color="9C0006",
                    )

                    cell.fill = PatternFill(
                        start_color="FFC7CE",
                        end_color="FFC7CE",
                        fill_type="solid",
                    )

                elif (
                    "Present" in status_value
                    or "P" in status_value
                ):

                    cell.font = Font(
                        name="Calibri",
                        size=11,
                        bold=True,
                        color="006100",
                    )

                    cell.fill = PatternFill(
                        start_color="C6EFCE",
                        end_color="C6EFCE",
                        fill_type="solid",
                    )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

            cell.border = thin_border

    # ========================================================
    # AUTO WIDTH
    # ========================================================

    for column in worksheet.columns:

        max_length = 0

        column_letter = get_column_letter(
            column[0].column
        )

        for cell in column:

            if cell.value is not None:

                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        worksheet.column_dimensions[
            column_letter
        ].width = max(
            max_length + 5,
            14
        )

    workbook.save(output)

    excel_data = output.getvalue()

    strlit.download_button(
        label="📥 تحميل تقرير Excel",
        data=excel_data,
        file_name=(
            f"Daily_Attendance_Report_"
            f"{selected_date_str}.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

    # ========================================================
    # 16. VIEW BUTTONS
    # ========================================================

    if is_today:

        if strlit.button(
            f"👥 كافة موظفي الشركة النشطين "
            f"({len(active_employees)})",
            use_container_width=True,
        ):

            strlit.session_state[
                "selected_view"
            ] = "all"

        column_present, column_late = strlit.columns(2)

        with column_present:

            if strlit.button(
                f"🟢 المتواجدون "
                f"({len(present_staff)})",
                use_container_width=True,
            ):

                strlit.session_state[
                    "selected_view"
                ] = "present"

        with column_late:

            if strlit.button(
                f"⏰ المتأخرون "
                f"({len(late_staff)})",
                use_container_width=True,
            ):

                strlit.session_state[
                    "selected_view"
                ] = "late"

        column_checkout, column_absent = strlit.columns(2)

        with column_checkout:

            if strlit.button(
                f"🏁 المنصرفون "
                f"({len(checkout_staff)})",
                use_container_width=True,
            ):

                strlit.session_state[
                    "selected_view"
                ] = "checkout"

        with column_absent:

            if strlit.button(
                f"❌ الغيابات "
                f"({len(absent_staff)})",
                use_container_width=True,
            ):

                strlit.session_state[
                    "selected_view"
                ] = "absent"

        column_leave, column_dummy = strlit.columns(2)

        with column_leave:

            if strlit.button(
                f"🏖️ الإجازات "
                f"({len(leave_staff)})",
                use_container_width=True,
            ):

                strlit.session_state[
                    "selected_view"
                ] = "leave"

    # ========================================================
    # 17. DEVICES
    # ========================================================

    with strlit.expander(
        "🖨️ أجهزة الحضور والانصراف المرتبطة",
        expanded=False,
    ):

        if devices:

            device_rows = []

            for device in devices:

                if not isinstance(device, dict):
                    continue

                device_name = (
                    device.get("alias")
                    or device.get("terminal_name")
                    or device.get("sn")
                    or "جهاز غير محدد"
                )

                device_sn = device.get(
                    "sn",
                    "N/A"
                )

                device_ip = device.get(
                    "ip_address",
                    "غير متوفر"
                )

                last_activity = device.get(
                    "last_activity"
                )

                status_badge = (
                    "<span class='badge-absent'>"
                    "غير متصل 🔴"
                    "</span>"
                )

                if last_activity:

                    try:

                        last_activity_dt = datetime.strptime(
                            str(last_activity)[:19],
                            "%Y-%m-%d %H:%M:%S"
                        )

                        now_naive = datetime.now()

                        seconds_since_activity = (
                            now_naive
                            - last_activity_dt
                        ).total_seconds()

                        if (
                            seconds_since_activity
                            < 1800
                        ):

                            status_badge = (
                                "<span "
                                "class='badge-present'>"
                                "متصل 🟢"
                                "</span>"
                            )

                    except Exception:
                        pass

                device_rows.append(
                    "<tr>"
                    f"<td>{clean_txt(device_name)}</td>"
                    f"<td>{clean_txt(device_sn)}</td>"
                    f"<td>{clean_txt(device_ip)}</td>"
                    f"<td>{status_badge}</td>"
                    "</tr>"
                )

            if device_rows:

                strlit.markdown(
                    f"""
                    <table
                        class="responsive-grid-table"
                    >
                        <tr>
                            <th>اسم الجهاز</th>
                            <th>الرقم التسلسلي (SN)</th>
                            <th>عنوان IP</th>
                            <th>الحالة</th>
                        </tr>

                        {"".join(device_rows)}

                    </table>
                    """,
                    unsafe_allow_html=True,
                )

        else:

            strlit.info(
                "لم يتم العثور على أجهزة BioTime."
            )

    # ========================================================
    # 18. SEARCH
    # ========================================================

    search_query = strlit.text_input(
        "",
        placeholder=TEXT_CONFIG[
            "search_placeholder"
        ],
        label_visibility="collapsed",
    ).strip().lower()

    def matches_search(code, name):

        if not search_query:
            return True

        return (
            search_query in str(code).lower()
            or search_query in str(name).lower()
        )

    view = strlit.session_state[
        "selected_view"
    ]

    # ========================================================
    # 19. ALL EMPLOYEES
    # ========================================================

    if view == "all":

        rows = []

        for code, employee_data in (
            active_employees.items()
        ):

            if not matches_search(
                code,
                employee_data["name"]
            ):
                continue

            rows.append(
                "<tr>"
                f"<td>{code}</td>"
                f"<td>{employee_data['name']}</td>"
                f"<td>{employee_data['dept']}</td>"
                "<td>"
                "<span class='badge-present'>"
                "نشط"
                "</span>"
                "</td>"
                "</tr>"
            )

        strlit.markdown(
            f"""
            <table
                class="responsive-grid-table"
            >
                <tr>
                    <th
                        colspan="4"
                        class="table-main-title-header"
                    >
                        {
                            TEXT_CONFIG[
                                "header_all"
                            ].format(
                                len(active_employees)
                            )
                        }
                    </th>
                </tr>

                <tr>
                    <th>الكود</th>
                    <th>الاسم</th>
                    <th>القسم</th>
                    <th>الحالة</th>
                </tr>

                {"".join(rows)}

            </table>
            """,
            unsafe_allow_html=True,
        )

    # ========================================================
    # 20. PRESENT
    # ========================================================

    elif view == "present":

        if present_staff:

            rows = []

            for code, name, department, time, device in present_staff:

                if not matches_search(
                    code,
                    name
                ):
                    continue

                rows.append(
                    "<tr>"
                    f"<td>{code}</td>"
                    f"<td>{name}</td>"
                    f"<td>{department}</td>"
                    f"<td>{time}</td>"
                    f"<td>{device}</td>"
                    "</tr>"
                )

            strlit.markdown(
                f"""
                <table
                    class="responsive-grid-table"
                >
                    <tr>
                        <th
                            colspan="5"
                            class="table-main-title-header"
                        >
                            {
                                TEXT_CONFIG[
                                    "header_present"
                                ].format(
                                    len(present_staff)
                                )
                            }
                        </th>
                    </tr>

                    <tr>
                        <th>الكود</th>
                        <th>الاسم</th>
                        <th>القسم</th>
                        <th>الدخول</th>
                        <th>جهاز البصمة</th>
                    </tr>

                    {"".join(rows)}

                </table>
                """,
                unsafe_allow_html=True,
            )

        else:

            strlit.info(
                "لا يوجد موظفون متواجدون حالياً."
            )

    # ========================================================
    # 21. LATE
    # ========================================================

    elif view == "late":

        if late_staff:

            rows = []

            for code, name, department, time, device in late_staff:

                if not matches_search(
                    code,
                    name
                ):
                    continue

                rows.append(
                    "<tr>"
                    f"<td>{code}</td>"
                    f"<td>{name}</td>"
                    f"<td>{department}</td>"
                    f"<td>{time}</td>"
                    "<td>"
                    "<span class='badge-late'>"
                    "متأخر"
                    "</span>"
                    "</td>"
                    f"<td>{device}</td>"
                    "</tr>"
                )

            strlit.markdown(
                f"""
                <table
                    class="responsive-grid-table"
                >
                    <tr>
                        <th
                            colspan="6"
                            class="table-main-title-header"
                        >
                            {
                                TEXT_CONFIG[
                                    "header_late"
                                ].format(
                                    len(late_staff)
                                )
                            }
                        </th>
                    </tr>

                    <tr>
                        <th>الكود</th>
                        <th>الاسم</th>
                        <th>القسم</th>
                        <th>الدخول</th>
                        <th>الحالة</th>
                        <th>جهاز البصمة</th>
                    </tr>

                    {"".join(rows)}

                </table>
                """,
                unsafe_allow_html=True,
            )

        else:

            strlit.info(
                "لا يوجد موظفون متأخرون."
            )

    # ========================================================
    # 22. CHECKOUT
    # ========================================================

    elif view == "checkout":

        if checkout_staff:

            rows = []

            for code, name, department, time, device in checkout_staff:

                if not matches_search(
                    code,
                    name
                ):
                    continue

                rows.append(
                    "<tr>"
                    f"<td>{code}</td>"
                    f"<td>{name}</td>"
                    f"<td>{department}</td>"
                    f"<td>{time}</td>"
                    f"<td>{device}</td>"
                    "</tr>"
                )

            strlit.markdown(
                f"""
                <table
                    class="responsive-grid-table"
                >
                    <tr>
                        <th
                            colspan="5"
                            class="table-main-title-header"
                        >
                            {
                                TEXT_CONFIG[
                                    "header_checkout"
                                ].format(
                                    len(checkout_staff)
                                )
                            }
                        </th>
                    </tr>

                    <tr>
                        <th>الكود</th>
                        <th>الاسم</th>
                        <th>القسم</th>
                        <th>الانصراف</th>
                        <th>جهاز البصمة</th>
                    </tr>

                    {"".join(rows)}

                </table>
                """,
                unsafe_allow_html=True,
            )

        else:

            strlit.info(
                "لا يوجد موظفون منصرفون."
            )

    # ========================================================
    # 23. LEAVE
    # ========================================================

    elif view == "leave":

        if leave_staff:

            rows = []

            for code, name, department, reason in leave_staff:

                if not matches_search(
                    code,
                    name
                ):
                    continue

                rows.append(
                    "<tr>"
                    f"<td>{code}</td>"
                    f"<td>{name}</td>"
                    f"<td>{department}</td>"
                    "<td>"
                    "<span class='badge-leave'>"
                    f"{reason}"
                    "</span>"
                    "</td>"
                    "</tr>"
                )

            strlit.markdown(
                f"""
                <table
                    class="responsive-grid-table"
                >
                    <tr>
                        <th
                            colspan="4"
                            class="table-main-title-header"
                        >
                            {
                                TEXT_CONFIG[
                                    "header_leave"
                                ].format(
                                    len(leave_staff)
                                )
                            }
                        </th>
                    </tr>

                    <tr>
                        <th>الكود</th>
                        <th>الاسم</th>
                        <th>القسم</th>
                        <th>نوع الإجازة</th>
                    </tr>

                    {"".join(rows)}

                </table>
                """,
                unsafe_allow_html=True,
            )

        else:

            strlit.info(
                "لا يوجد موظفون في إجازة."
            )

    # ========================================================
    # 24. ABSENT
    # ========================================================

    elif view == "absent":

        if absent_staff:

            rows = []

            for code, name, department in absent_staff:

                if not matches_search(
                    code,
                    name
                ):
                    continue

                rows.append(
                    "<tr>"
                    f"<td>{code}</td>"
                    f"<td>{name}</td>"
                    f"<td>{department}</td>"
                    "<td>"
                    "<span class='badge-absent'>"
                    "غياب"
                    "</span>"
                    "</td>"
                    "</tr>"
                )

            strlit.markdown(
                f"""
                <table
                    class="responsive-grid-table"
                >
                    <tr>
                        <th
                            colspan="4"
                            class="table-main-title-header"
                        >
                            {
                                TEXT_CONFIG[
                                    "header_absent"
                                ].format(
                                    len(absent_staff)
                                )
                            }
                        </th>
                    </tr>

                    <tr>
                        <th>الكود</th>
                        <th>الاسم</th>
                        <th>القسم</th>
                        <th>الحالة</th>
                    </tr>

                    {"".join(rows)}

                </table>
                """,
                unsafe_allow_html=True,
            )

        else:

            strlit.info(
                "لا يوجد غياب."
            )


# ============================================================
# 25. ERROR HANDLING
# ============================================================

except Exception as error:

    error_message = str(error)

    strlit.markdown(
        f"""
        <div class="api-error-box">

            <strong>
                ❌ BioTime API Error
            </strong>

            <br><br>

            {clean_txt(error_message)}

            <br><br>

            <small>
                The password and authentication token are
                intentionally not displayed.
            </small>

        </div>
        """,
        unsafe_allow_html=True,
    )
