import streamlit as st
import pandas as pd
from datetime import datetime
import re
import io
import requests
import base64
import urllib.parse
import os
from streamlit_gsheets import GSheetsConnection

# ==========================================
# SYSTEM CONFIGURATION & API
# ==========================================
def get_runtime_secret(name):
    """Load deployment secrets without committing them to the public repository."""
    try:
        value = st.secrets.get(name, "")
        if value:
            return str(value).strip()
    except Exception:
        pass
    return str(os.environ.get(name, "")).strip()


IMGBB_API_KEY = get_runtime_secret("IMGBB_API_KEY")

st.set_page_config(page_title="Al-Qasr Al-Zahabi ERP", layout="wide", page_icon="🏢")

query_params = st.query_params
is_tv_mode = "tv" in query_params or query_params.get("mode") == "tv"

if is_tv_mode:
    st.markdown("""
        <style>
            [data-testid='stSidebar'] {display: none !important;}
            header {visibility: hidden !important;}
            .stApp { background-color: #f8f9fa; direction: rtl; text-align: right; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
            h1, h2, h3, h4, p, span, label, div { text-align: right; }
        </style>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
        <style>
            .stApp { background-color: #f8f9fa; direction: rtl; text-align: right; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
            h1, h2, h3, h4, p, span, label, div { text-align: right; }
            .stTabs [data-baseweb="tab-list"] { gap: 10px; border-bottom: 2px solid #e2e8f0; }
            .stTabs [data-baseweb="tab"] { background-color: transparent; border-radius: 4px 4px 0 0; padding: 10px 20px; font-weight: 600; color: #4a5568; }
            .stTabs [aria-selected="true"] { border-bottom: 3px solid #3182ce; color: #2b6cb0; background-color: #ebf8ff; }
            .erp-card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin-bottom: 20px; }
            .locked-card { background: #fff5f5; padding: 20px; border: 1px solid #feb2b2; border-radius: 8px; margin-bottom: 20px; }
            .invoice-box { background: white; padding: 30px; border: 1px solid #e2e8f0; border-radius: 8px; max-width: 800px; margin: auto; }
            .invoice-header { text-align: center; border-bottom: 2px solid #2b6cb0; padding-bottom: 15px; margin-bottom: 20px; }
        </style>
    """, unsafe_allow_html=True)

# ==========================================
# DATABASE ORM (DocType Engine)
# ==========================================
conn = st.connection("gsheets", type=GSheetsConnection)

CASE_STATUS_OPEN = "مفتوح"
CASE_STATUS_CLOSED = "مغلق"
COLLECTION_NOT_READY = "لم يجهز للتسليم بعد"
COLLECTION_AWAITING = "بانتظار تأكيد الاستلام"
COLLECTION_PAID_AWAITING_CLOSE = "تم تسجيل القبض - بانتظار إغلاق الحالة"  # legacy value
COLLECTION_SPECIAL_AWAITING = "حالة خاصة - بانتظار الاستلام"
COLLECTION_CLOSED = "تم الاستلام وإغلاق الحالة"

# Terminal no-charge notes. Customer S/D cases wait for a manual collection
# confirmation; partner V cases close automatically because no payment is due.
NO_CHARGE_SPECIAL_CASES = {
    "يعمل من المصدر",
    "الزبون رفض الإصلاح",
    "مكلف",
    "كفالة",
    "غير قابل للإصلاح",
    "لا يوجد عطل",
}

SCHEMA = {
    "Ledger": [
        "service_id", "tool_name", "customer_name", "phone_number", "warranty_status", "document_origin",
        "reported_issue", "technician", "status", "cost_debit", "payment_credit", "balance",
        "spare_parts", "resolution_notes", "remarks", "date_logged", "date_resolved",
        "accessories", "loaner_item", "priority", "tool_photo_link",
        "source_account", "source_account_g", "source_document_count", "document_history",
        "repair_stage", "collection_status", "special_case", "partner_claim_status", "partner_claim_amount",
        "case_status", "closed_at", "closed_by", "close_note"
    ],
    "Stock": ["item_code", "item_name", "quantity", "price"],
    "Hawara": ["order_id", "order_type", "linked_service_id", "courier", "delivery_note", "document_link", "status", "date_logged"],
    "Dispatch": ["dispatch_id", "service_id", "customer_name", "courier", "delivery_note", "document_link", "date"]
}

STOCK_COLUMNS = SCHEMA["Stock"]


def normalize_doc_string(val):
    return re.sub(r'\s+', ' ', str(val or '')).strip()


def get_status_rank(val):
    """Ranking used only for source-document stage comparison."""
    s = normalize_doc_string(val)
    if "قبض" in s or "Collected" in s:
        return 4
    if "خ صيانة" in s or "حساب وكيل" in s:
        return 3
    if "مبيع خ ص" in s or "جاهز" in s:
        return 2
    if "اد خ ص" in s or "المعالجة" in s:
        return 1
    return 0


def map_document_to_status(doc_string, cost=0.0, collection_status="", special_case="", partner_claim_status="", case_status=""):
    doc = normalize_doc_string(doc_string)
    special = normalize_doc_string(special_case)
    collection = normalize_doc_string(collection_status)
    partner = normalize_doc_string(partner_claim_status)
    case = normalize_doc_string(case_status)

    if case == CASE_STATUS_CLOSED or COLLECTION_CLOSED in collection or collection == "تم التحصيل والتسليم":
        return "مغلق - تم الاستلام (Closed)"
    if special:
        return f"حالة خاصة - بانتظار الاستلام ({special})"
    if "خ صيانة" in doc and "بانتظار" not in partner:
        return "مغلق محاسبياً - حساب شريك (Partner Claimed)"
    if "قبض" in doc:
        return "مغلق - تم الاستلام (Closed)"
    if "مبيع خ ص" in doc or "جاهز" in doc:
        return "جاهز للتسليم - بانتظار الاستلام (Ready / Awaiting Collection)"
    if "اد خ ص" in doc or "المعالجة" in doc:
        return "قيد المعالجة (In Progress)"
    return "قيد الانتظار (Waiting)"


def parse_excel_date(value):
    try:
        if value is None or str(value).strip() in ("", "nan", "NaT"):
            return ""
        if isinstance(value, (int, float)) and not pd.isna(value):
            return (pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")).strftime("%Y-%m-%d")
        parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
        return parsed.strftime("%Y-%m-%d") if pd.notna(parsed) else ""
    except Exception:
        return ""


def extract_service_id(account_value, account_g_value=""):
    text = f"{account_g_value} {account_value}"
    m = re.search(r'\b([SDV]\d+)\b', text, re.IGNORECASE)
    return m.group(1).upper() if m else ""


def clean_text(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).replace("\n", " ").strip()


def parse_account_details(account_text, service_id=""):
    """Best-effort parsing of Ameen account description without assuming a fixed customer position."""
    text = clean_text(account_text)
    if not text:
        return "", "", "", "", ""
    parts = [p.strip() for p in re.split(r'[-–—]+', text) if p.strip()]
    if service_id and parts and parts[0].upper() == service_id.upper():
        parts = parts[1:]

    phone = ""
    phone_idx = None
    for i, p in enumerate(parts):
        digits = re.sub(r'\D', '', p)
        if 8 <= len(digits) <= 15:
            phone = digits
            phone_idx = i
            break

    known_code_idx = None
    for i, p in enumerate(parts):
        if re.fullmatch(r'[A-Za-z]{2,}\d+[A-Za-z0-9]*', p) or re.fullmatch(r'[A-Za-z0-9]{5,}', p):
            if re.search(r'[A-Za-z]', p) and re.search(r'\d', p):
                known_code_idx = i
                break

    warranty = "ضمن كفالة" if any(k in text for k in ["كفالة", "ضمان", "مجاني"]) else "خارج الكفالة"

    item = ""
    customer = ""
    issue = ""
    if known_code_idx is not None:
        item_parts = parts[:known_code_idx]
        if item_parts:
            if len(item_parts) >= 2:
                customer = item_parts[-1]
                item = " - ".join(item_parts[:-1])
            else:
                item = item_parts[0]
        tail_start = known_code_idx + 1
        tail = parts[tail_start:]
        if phone_idx is not None and phone_idx >= tail_start:
            if phone_idx + 1 < len(parts):
                issue = " - ".join(parts[phone_idx + 1:])
        elif tail:
            issue = " - ".join(tail)
    else:
        pre = parts[:phone_idx] if phone_idx is not None else parts
        if len(pre) >= 2:
            item = pre[0]
            customer = " - ".join(pre[1:])
        elif len(pre) == 1:
            item = pre[0]
        if phone_idx is not None and phone_idx + 1 < len(parts):
            issue = " - ".join(parts[phone_idx + 1:])

    issue = issue or text
    return item or "غير محدد", customer or "غير محدد", phone, warranty, issue


def special_case_from_remarks(text):
    t = normalize_doc_string(text)
    if any(k in t for k in ["رفض الإصلاح", "رفض الاصلاح", "الزبون رفض", "رفض الصيانة", "رفض التصليح", "لم يوافق على الإصلاح", "لم يوافق على الاصلاح"]):
        return "الزبون رفض الإصلاح"
    if any(k in t for k in ["يعمل من المصدر", "يعمل على المصدر"]):
        return "يعمل من المصدر"
    if "مكلف" in t:
        return "مكلف"
    if any(k in t for k in ["كفالة", "ضمان"]) and not any(k in t for k in ["خارج الكفالة", "خارج كفالة", "خارج الضمان"]):
        return "كفالة"
    if any(k in t for k in ["غير قابل للإصلاح", "غير قابل للاصلاح", "غير قابل للصيانة", "لاتصلح", "لا تصلح", "لا يمكن إصلاح", "لا يمكن اصلاح"]):
        return "غير قابل للإصلاح"
    if "لايوجد عطل" in t or "لا يوجد عطل" in t:
        return "لا يوجد عطل"
    return ""


def is_no_charge_special_case(value):
    return normalize_doc_string(value) in NO_CHARGE_SPECIAL_CASES


def is_zero_amount(value, tolerance=0.000001):
    """True only for a numeric value that was explicitly parsed as zero."""
    try:
        number = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
        return pd.notna(number) and abs(float(number)) <= tolerance
    except Exception:
        return False


def normalize_ameen_dataframe(raw_excel):
    """Turn Ameen's A:L ledger into one row per service ticket while retaining document history."""
    df = raw_excel.copy()
    if df.empty:
        return pd.DataFrame()

    # Ameen export is fixed A:L. Headers are not unique ("مدين"/"دائن" repeat),
    # therefore we deliberately map by position rather than by header text.
    expected = ["account", "prev_balance", "debit", "credit", "uncolllected", "current_balance", "account_name", "date", "origin", "debit_line", "credit_line", "statement"]
    if df.shape[1] >= 12:
        df = df.iloc[:, :12].copy()
        df.columns = expected
    else:
        return pd.DataFrame()

    # Drop header rows repeated inside an export.
    df = df[df["origin"].astype(str).str.strip().str.lower() != "أصل السند"].copy()
    df["service_id"] = [extract_service_id(a, g) for a, g in zip(df["account"], df["account_name"])]
    df = df[df["service_id"].str.strip() != ""].copy()
    if df.empty:
        return pd.DataFrame()

    df["date_parsed"] = df["date"].apply(parse_excel_date)
    for c in ["debit_line", "credit_line", "debit", "credit"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df["origin_clean"] = df["origin"].apply(normalize_doc_string)
    df["statement_clean"] = df["statement"].apply(clean_text)
    df["source_account"] = df["account"].apply(clean_text)
    df["source_account_g"] = df["account_name"].apply(clean_text)

    records = []
    for sid, grp in df.groupby("service_id", sort=False):
        # Preserve source order, then use dates to identify the latest row.
        grp = grp.copy()
        grp["sort_date"] = pd.to_datetime(grp["date_parsed"], errors="coerce")
        grp["origin_rank"] = grp["origin_clean"].apply(get_status_rank)
        grp = grp.sort_values(["sort_date", "origin_rank"]).reset_index(drop=True)

        header_text = next((x for x in grp["source_account_g"] if x), "") or next((x for x in grp["source_account"] if x), "")
        item, customer, phone, warranty, issue = parse_account_details(header_text, sid)
        all_statements = [x for x in grp["statement_clean"].tolist() if x and x not in ("13", "nan")]
        special_cases = [special_case_from_remarks(x) for x in all_statements]
        special_case = next((x for x in special_cases if x), "")
        has_entry = grp["origin_clean"].str.contains("اد خ ص", na=False).any()
        has_sale = grp["origin_clean"].str.contains("مبيع خ ص", na=False).any()
        has_collect = grp["origin_clean"].str.contains("قبض", na=False).any()
        has_partner_claim = grp["origin_clean"].str.contains("خ صيانة", na=False).any()
        is_partner = sid.upper().startswith("V")

        sale_mask = grp["origin_clean"].str.contains("مبيع خ ص", na=False)
        cost_debit = float(grp.loc[sale_mask, "debit_line"].sum())
        payment_credit = float(grp.loc[grp["origin_clean"].str.contains("قبض", na=False), "credit_line"].sum())
        partner_claim_amount = float(grp.loc[grp["origin_clean"].str.contains("خ صيانة", na=False), "credit_line"].sum())

        # New temporary business rule (2026-09):
        # For customer S/D cases, a real Ameen sales invoice (مبيع خ ص) whose
        # total amount is exactly zero is terminal. No separate collection
        # confirmation is required. V partner cases still follow partner-claim rules.
        zero_value_customer_sale = (
            not is_partner
            and has_sale
            and is_zero_amount(cost_debit)
        )

        auto_close_reason = ""
        if has_collect:
            # Business rule: a قبض entry means the customer physically collected the tool.
            auto_close_reason = "قبض في كشف الأمين - تم الاستلام"
        elif zero_value_customer_sale:
            auto_close_reason = "فاتورة مبيع خ ص بقيمة صفر - إغلاق تلقائي"
        elif is_partner and has_partner_claim:
            auto_close_reason = "إغلاق شريك - خ صيانة"
        elif is_partner and is_no_charge_special_case(special_case):
            auto_close_reason = f"إغلاق شريك - حالة خاصة: {special_case}"

        case_status = CASE_STATUS_CLOSED if auto_close_reason else CASE_STATUS_OPEN

        if case_status == CASE_STATUS_CLOSED:
            collection_status = COLLECTION_CLOSED
        elif special_case:
            collection_status = COLLECTION_SPECIAL_AWAITING
        elif has_sale:
            collection_status = COLLECTION_AWAITING
        else:
            collection_status = COLLECTION_NOT_READY

        if not is_partner:
            partner_claim_status = "غير منطبق"
        elif has_partner_claim:
            partner_claim_status = "تمت مطالبة الشريك / تم التحصيل"
        elif is_no_charge_special_case(special_case):
            partner_claim_status = "لا مطالبة - حالة خاصة"
        elif has_sale or has_entry:
            partner_claim_status = "بانتظار مطالبة الشريك"
        else:
            partner_claim_status = "غير محدد"

        if case_status == CASE_STATUS_CLOSED:
            repair_stage = CASE_STATUS_CLOSED
        elif special_case:
            repair_stage = "جاهز للتسليم"
        elif has_sale:
            repair_stage = "جاهز للتسليم"
        elif has_entry:
            repair_stage = "قيد المعالجة"
        else:
            repair_stage = "قيد الانتظار"

        latest = grp.iloc[-1]
        date_logged = next((x for x in grp["date_parsed"].tolist() if x), datetime.now().strftime("%Y-%m-%d"))

        def latest_event_date(mask=None):
            date_values = grp.loc[mask, "date_parsed"].tolist() if mask is not None else grp["date_parsed"].tolist()
            valid_dates = [x for x in date_values if x]
            return valid_dates[-1] if valid_dates else datetime.now().strftime("%Y-%m-%d")

        closure_date = ""
        if auto_close_reason:
            if has_collect:
                closure_date = latest_event_date(grp["origin_clean"].str.contains("قبض", na=False))
            elif zero_value_customer_sale:
                closure_date = latest_event_date(sale_mask)
            elif is_partner and has_partner_claim:
                closure_date = latest_event_date(grp["origin_clean"].str.contains("خ صيانة", na=False))
            else:
                closure_date = latest_event_date()
        date_resolved = closure_date

        # The source remarks are preserved verbatim; don't assign meanings to unconfirmed document types.
        doc_history = " | ".join(f"{d} :: {o}" for d, o in zip(grp["date_parsed"], grp["origin_clean"]) if o)
        remarks = " | ".join(dict.fromkeys(all_statements))
        latest_origin = latest["origin_clean"]
        status = map_document_to_status(latest_origin, cost_debit, collection_status, special_case, partner_claim_status, case_status)

        records.append({
            "service_id": sid, "tool_name": item, "customer_name": customer, "phone_number": phone,
            "warranty_status": warranty, "document_origin": latest_origin, "reported_issue": issue,
            "technician": "Ameen Import", "status": status, "cost_debit": cost_debit, "payment_credit": payment_credit,
            "balance": max(cost_debit - payment_credit, 0.0), "spare_parts": "لا حاجة / متوفرة", "resolution_notes": "",
            "remarks": remarks, "date_logged": date_logged, "date_resolved": date_resolved,
            "accessories": "", "loaner_item": "", "priority": "عادي", "tool_photo_link": "",
            "source_account": next((x for x in grp["source_account"] if x), ""),
            "source_account_g": next((x for x in grp["source_account_g"] if x), ""),
            "source_document_count": int(len(grp)), "document_history": doc_history,
            "repair_stage": repair_stage, "collection_status": collection_status,
            "special_case": special_case, "partner_claim_status": partner_claim_status,
            "partner_claim_amount": partner_claim_amount,
            "case_status": case_status,
            "closed_at": closure_date,
            "closed_by": "Ameen Import" if auto_close_reason else "",
            "close_note": auto_close_reason,
        })
    return pd.DataFrame(records)


def apply_workflow_columns(df):
    """Backfill and enforce workflow rules for existing Google Sheet rows."""
    if df.empty:
        return df
    for c, default in {
        "source_account": "", "source_account_g": "", "source_document_count": 1, "document_history": "",
        "repair_stage": "", "collection_status": "", "special_case": "", "partner_claim_status": "غير منطبق", "partner_claim_amount": 0.0,
        "case_status": CASE_STATUS_OPEN, "closed_at": "", "closed_by": "", "close_note": ""
    }.items():
        if c not in df.columns:
            df[c] = default
    df["source_document_count"] = pd.to_numeric(df["source_document_count"], errors="coerce").fillna(1).astype(int)
    df["partner_claim_amount"] = pd.to_numeric(df["partner_claim_amount"], errors="coerce").fillna(0.0)

    for idx, r in df.iterrows():
        doc = normalize_doc_string(r.get("document_origin", ""))
        history = normalize_doc_string(r.get("document_history", ""))
        sid = str(r.get("service_id", ""))
        is_partner = sid.upper().startswith("V")
        case_status = normalize_doc_string(r.get("case_status", "")) or CASE_STATUS_OPEN

        legacy_closed = (
            normalize_doc_string(r.get("collection_status", "")) == "تم التحصيل والتسليم"
            or "تم التحصيل والتسليم" in normalize_doc_string(r.get("status", ""))
        )
        if legacy_closed:
            case_status = CASE_STATUS_CLOSED
            if not r.get("closed_at"):
                df.at[idx, "closed_at"] = str(r.get("date_resolved", "") or r.get("date_logged", ""))
            if not r.get("closed_by"):
                df.at[idx, "closed_by"] = "ترحيل النظام"

        special = normalize_doc_string(r.get("special_case", "")) or special_case_from_remarks(r.get("remarks", ""))
        df.at[idx, "special_case"] = special

        has_collect_signal = "قبض" in doc or "قبض" in history
        has_sale_signal = "مبيع خ ص" in doc or "مبيع خ ص" in history
        has_entry_signal = "اد خ ص" in doc or "اد خ ص" in history
        has_partner_claim_signal = is_partner and ("خ صيانة" in doc or "خ صيانة" in history)

        # Guard the zero-invoice rule so a manually-created Ready row with a blank/zero
        # cost does not get closed by mistake. It applies only to Ameen-imported rows.
        is_ameen_import = normalize_doc_string(r.get("technician", "")) == "Ameen Import"
        zero_value_customer_sale = (
            not is_partner
            and is_ameen_import
            and has_sale_signal
            and is_zero_amount(r.get("cost_debit", None))
        )

        auto_close_reason = ""
        if case_status != CASE_STATUS_CLOSED:
            if has_collect_signal:
                auto_close_reason = "قبض في كشف الأمين - تم الاستلام"
            elif zero_value_customer_sale:
                auto_close_reason = "فاتورة مبيع خ ص بقيمة صفر - إغلاق تلقائي"
            elif has_partner_claim_signal:
                auto_close_reason = "إغلاق شريك - خ صيانة"
            elif is_partner and is_no_charge_special_case(special):
                auto_close_reason = f"إغلاق شريك - حالة خاصة: {special}"

        if auto_close_reason:
            case_status = CASE_STATUS_CLOSED
            if not normalize_doc_string(r.get("closed_at", "")):
                df.at[idx, "closed_at"] = str(r.get("date_resolved", "") or r.get("date_logged", "") or datetime.now().strftime("%Y-%m-%d"))
            if not normalize_doc_string(r.get("closed_by", "")):
                df.at[idx, "closed_by"] = "Ameen Import"
            if not normalize_doc_string(r.get("close_note", "")):
                df.at[idx, "close_note"] = auto_close_reason
            if not normalize_doc_string(r.get("date_resolved", "")):
                df.at[idx, "date_resolved"] = str(df.at[idx, "closed_at"]).split(" ")[0]

        df.at[idx, "case_status"] = case_status

        if case_status == CASE_STATUS_CLOSED:
            df.at[idx, "collection_status"] = COLLECTION_CLOSED
            df.at[idx, "repair_stage"] = CASE_STATUS_CLOSED
        else:
            if special:
                df.at[idx, "collection_status"] = COLLECTION_SPECIAL_AWAITING
                df.at[idx, "repair_stage"] = "جاهز للتسليم"
            elif has_sale_signal:
                df.at[idx, "collection_status"] = COLLECTION_AWAITING
                df.at[idx, "repair_stage"] = "جاهز للتسليم"
            elif has_entry_signal:
                df.at[idx, "collection_status"] = COLLECTION_NOT_READY
                df.at[idx, "repair_stage"] = "قيد المعالجة"
            elif not normalize_doc_string(r.get("collection_status", "")):
                df.at[idx, "collection_status"] = COLLECTION_NOT_READY
                df.at[idx, "repair_stage"] = "قيد الانتظار"

        if is_partner:
            if has_partner_claim_signal:
                df.at[idx, "partner_claim_status"] = "تمت مطالبة الشريك / تم التحصيل"
            elif is_no_charge_special_case(special):
                df.at[idx, "partner_claim_status"] = "لا مطالبة - حالة خاصة"
            elif not normalize_doc_string(r.get("partner_claim_status", "")) or normalize_doc_string(r.get("partner_claim_status", "")) == "غير منطبق":
                df.at[idx, "partner_claim_status"] = "بانتظار مطالبة الشريك"
        else:
            df.at[idx, "partner_claim_status"] = "غير منطبق"

        df.at[idx, "status"] = map_document_to_status(
            doc,
            float(r.get("cost_debit", 0) or 0),
            df.at[idx, "collection_status"],
            df.at[idx, "special_case"],
            df.at[idx, "partner_claim_status"],
            case_status,
        )
    return df


def preserve_manual_closures(imported_df, existing_df):
    """Keep app-managed closure fields when a later Ameen statement is imported."""
    if imported_df is None or imported_df.empty or existing_df is None or existing_df.empty:
        return imported_df

    result = imported_df.copy()
    existing = existing_df.copy()
    for col, default in {
        "case_status": CASE_STATUS_OPEN, "closed_at": "", "closed_by": "", "close_note": ""
    }.items():
        if col not in existing.columns:
            existing[col] = default
        if col not in result.columns:
            result[col] = default

    closed = existing[existing["case_status"].astype(str).eq(CASE_STATUS_CLOSED)].copy()
    if closed.empty:
        return result
    closed = closed.drop_duplicates("service_id", keep="last").set_index("service_id")

    for idx, row in result.iterrows():
        sid = str(row.get("service_id", ""))
        if sid not in closed.index:
            continue
        previous = closed.loc[sid]
        for col in ("case_status", "closed_at", "closed_by", "close_note"):
            result.at[idx, col] = previous.get(col, "")
        result.at[idx, "collection_status"] = COLLECTION_CLOSED
        result.at[idx, "repair_stage"] = CASE_STATUS_CLOSED
        result.at[idx, "status"] = "مغلق - تم الاستلام (Closed)"
        if previous.get("closed_at", ""):
            result.at[idx, "date_resolved"] = str(previous.get("closed_at", "")).split(" ")[0]
    return result


def deduplicate_ledger(df):
    """Keep one row per service case while preserving the furthest known workflow stage."""
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df

    work = df.copy()
    if "service_id" not in work.columns:
        return work

    work["service_id"] = work["service_id"].astype(str).replace({"nan": "", "None": ""})
    work = work[work["service_id"].str.strip() != ""].copy()
    if work.empty:
        return work

    # A repeated service reference represents the same repair case.
    # Prefer the highest workflow stage, then the newest known date.
    if "document_origin" in work.columns:
        work["_workflow_rank"] = work["document_origin"].apply(get_status_rank)
    elif "repair_stage" in work.columns:
        stage_rank = {
            "قيد الانتظار": 0,
            "قيد المعالجة": 1,
            "جاهز للتسليم": 2,
            "تم التحصيل والتسليم": 4,
        }
        work["_workflow_rank"] = work["repair_stage"].map(stage_rank).fillna(0)
    else:
        work["_workflow_rank"] = 0
    if "case_status" in work.columns:
        work.loc[work["case_status"].astype(str).eq(CASE_STATUS_CLOSED), "_workflow_rank"] = 5

    date_series = pd.Series(pd.NaT, index=work.index, dtype="datetime64[ns]")
    for col in ("date_resolved", "date_logged"):
        if col in work.columns:
            parsed = pd.to_datetime(work[col], dayfirst=True, errors="coerce")
            date_series = date_series.fillna(parsed)
    work["_workflow_date"] = date_series

    work = work.sort_values(
        ["service_id", "_workflow_rank", "_workflow_date"],
        ascending=[True, True, True],
        kind="stable",
        na_position="first",
    )
    work = work.groupby("service_id", as_index=False, sort=False).tail(1)
    return work.drop(columns=["_workflow_rank", "_workflow_date"], errors="ignore").reset_index(drop=True)


def get_doctype(doctype_name):
    try:
        df = conn.read(worksheet=doctype_name, ttl=0)
        df = df.dropna(how='all')

        for col in SCHEMA[doctype_name]:
            if col not in df.columns: df[col] = ""

        if doctype_name == "Ledger":
            df['service_id'] = df['service_id'].astype(str).replace({'nan': '', 'None': ''})
            df = df[df['service_id'].str.strip() != ""]

            for col in SCHEMA["Ledger"]:
                if col not in ['cost_debit', 'payment_credit', 'balance']:
                    df[col] = df[col].fillna("").astype(str).replace({'nan': '', 'None': ''})

            for col in ['cost_debit', 'payment_credit', 'balance']:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

            df.loc[df['spare_parts'] == "", 'spare_parts'] = "لا حاجة / متوفرة"
            df = apply_workflow_columns(df)
            return deduplicate_ledger(df)

        return df
    except Exception as e:
        error_msg = str(e).strip()
        if error_msg != doctype_name and "not found" not in error_msg.lower() and "HTTPError" in error_msg:
            st.error(f"⚠️ خطأ في الاتصال (Connection Error): {error_msg}")
        return pd.DataFrame(columns=SCHEMA[doctype_name])


def save_doctype(doctype_name, df):
    if doctype_name == "Ledger":
        df = apply_workflow_columns(df)
        df = deduplicate_ledger(df)
        df = df[df['service_id'].astype(str).str.strip() != ""]
    conn.update(worksheet=doctype_name, data=df)


def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='Export')
    return output.getvalue()


def upload_to_cloud(file_buffer):
    if not IMGBB_API_KEY:
        st.warning("⚠️ رفع الصور متوقف حتى يتم ضبط IMGBB_API_KEY في إعدادات التطبيق الآمنة.")
        return ""
    try:
        b64_img = base64.b64encode(file_buffer.getvalue()).decode("utf-8")
        res = requests.post(f"https://api.imgbb.com/1/upload?key={IMGBB_API_KEY}", data={"image": b64_img})
        if res.status_code == 200: return res.json()["data"]["url"]
    except: return ""
    return ""


def generate_next_id(branch_code, df):
    if df.empty: return f"{branch_code}1"
    branch_records = df[df['service_id'].astype(str).str.startswith(branch_code, na=False)]
    if branch_records.empty: return f"{branch_code}1"

    max_num = 0
    for sid in branch_records['service_id']:
        num_part = re.sub(r'\D', '', str(sid))
        if num_part: max_num = max(max_num, int(num_part))
    return f"{branch_code}{max_num + 1}"

# ==========================================
# AUTHENTICATION & WORKSPACE ROUTING
# ==========================================
if 'logged_in_user' not in st.session_state: st.session_state['logged_in_user'] = None
if 'current_module' not in st.session_state: st.session_state['current_module'] = 'Workspace'

if is_tv_mode:
    st.session_state['logged_in_user'] = "TV_Guest"
    st.session_state['current_module'] = 'TV_Display'

USERS = {}
admin_password = get_runtime_secret("APP_ADMIN_PASSWORD")
tech_password = get_runtime_secret("APP_TECH_PASSWORD")
if admin_password:
    USERS["admin"] = {"pass": admin_password, "role": "System Administrator"}
if tech_password:
    USERS["tech"] = {"pass": tech_password, "role": "Support Agent"}

if not is_tv_mode and not USERS:
    st.error("⚠️ يجب ضبط APP_ADMIN_PASSWORD أو APP_TECH_PASSWORD في إعدادات التطبيق الآمنة قبل تسجيل الدخول.")
    st.stop()

if st.session_state['logged_in_user'] is None:
    st.markdown("<div class='erp-card' style='max-width: 400px; margin: 100px auto; text-align: center;'>", unsafe_allow_html=True)
    st.title("🏢 ERP Login")
    st.subheader("القصر الذهبي للمعدات")
    with st.form("login_form"):
        u_in = st.text_input("اسم المستخدم (Username)")
        p_in = st.text_input("كلمة المرور (Password)", type="password")
        if st.form_submit_button("تسجيل الدخول", use_container_width=True):
            if u_in in USERS and USERS[u_in]["pass"] == p_in:
                st.session_state['logged_in_user'] = u_in
                st.rerun()
            else: st.error("بيانات الدخول غير صحيحة.")
    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

current_user = st.session_state['logged_in_user']
is_admin = current_user in USERS and "Administrator" in USERS[current_user]["role"]

ledger_df = get_doctype("Ledger")
stock_df = get_doctype("Stock")
hawara_df = get_doctype("Hawara")
dispatch_df = get_doctype("Dispatch")

stock_list = stock_df['item_name'].dropna().unique().tolist() if not stock_df.empty else []

if not is_tv_mode:
    with st.sidebar:
        st.title("🏢 ERPNext Workspace")
        st.markdown(f"**المستخدم:** {current_user} <br> **الدور:** {USERS.get(current_user, {}).get('role', 'Viewer')}", unsafe_allow_html=True)
        st.divider()

        st.caption("العمليات الأساسية (CORE MODULES)")
        if st.button("🏠 مساحة العمل (Workspace)", use_container_width=True): st.session_state['current_module'] = 'Workspace'
        if st.button("📺 شاشة الورشة (TV Display)", use_container_width=True): st.session_state['current_module'] = 'TV_Display'
        if st.button("🛠️ الدعم والصيانة (Support)", use_container_width=True): st.session_state['current_module'] = 'Support'
        if st.button("📦 المخزون (Stock)", use_container_width=True): st.session_state['current_module'] = 'Stock'
        if st.button("🚚 اللوجستيات (Logistics)", use_container_width=True): st.session_state['current_module'] = 'Logistics'

        st.caption("المالية والتقارير (ACCOUNTING)")
        if st.button("💰 المحاسبة (Accounting)", use_container_width=True): st.session_state['current_module'] = 'Accounting'

        st.divider()
        if st.button("🚪 تسجيل الخروج (Logout)", use_container_width=True):
            st.session_state['logged_in_user'] = None
            st.rerun()

if st.session_state['current_module'] == 'Workspace':
    st.title("مساحة العمل الموحدة (Workspace)")
    active_count = len(ledger_df[~ledger_df['case_status'].astype(str).eq(CASE_STATUS_CLOSED)]) if not ledger_df.empty else 0
    ready_count = len(ledger_df[ledger_df['status'].str.contains('جاهز', na=False)]) if not ledger_df.empty else 0
    total_rev = float(ledger_df['cost_debit'].sum()) if not ledger_df.empty else 0.0

    col1, col2, col3 = st.columns(3)
    with col1: st.markdown(f"<div class='erp-card'><h3>🛠️ صيانة مفتوحة</h3><h1>{active_count}</h1></div>", unsafe_allow_html=True)
    with col2: st.markdown(f"<div class='erp-card'><h3>✅ أجهزة جاهزة للتسليم</h3><h1>{ready_count}</h1></div>", unsafe_allow_html=True)
    with col3: st.markdown(f"<div class='erp-card'><h3>💰 إجمالي المبيعات</h3><h1>${total_rev:,.2f}</h1></div>", unsafe_allow_html=True)

# ==========================================
# MODULE 2: TV WORKSHOP DISPLAY (KIOSK MODE)
# ==========================================
elif st.session_state['current_module'] == 'TV_Display':

    if is_tv_mode:
        # Smooth auto-scrolling engine (Voice completely removed)
        html_injection = """
        <script>
            let goingDown = true;
            const scrollSpeed = 1;
            const intervalTime = 30;

            let scrollInterval = setInterval(() => {
                if (goingDown) {
                    window.parent.scrollBy(0, scrollSpeed);
                    if ((window.parent.innerHeight + window.parent.scrollY) >= window.parent.document.body.offsetHeight - 5) {
                        goingDown = false;
                        setTimeout(() => {}, 2000);
                    }
                } else {
                    window.parent.scrollBy(0, -scrollSpeed);
                    if (window.parent.scrollY <= 0) {
                        goingDown = true;
                        setTimeout(() => {}, 2000);
                    }
                }
            }, intervalTime);

            setInterval(() => {
                fetch(window.parent.location.href)
                    .then(res => res.text())
                    .then(html => {}).catch(err => {});
            }, 30000);
        </script>
        """
        st.components.v1.html(html_injection, height=0)

    st.markdown("""
        <style>
            .tv-card-urgent { background: #ffe5e5; border-right: 15px solid #e53e3e; padding: 25px; border-radius: 10px; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
            .tv-card-delayed { background: #fffaf0; border-right: 15px solid #dd6b20; padding: 25px; border-radius: 10px; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
            .tv-card-normal { background: #ebf8ff; border-right: 15px solid #3182ce; padding: 25px; border-radius: 10px; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
            .tv-title { font-size: 32px; font-weight: bold; color: #1a202c; margin-bottom: 10px; }
            .tv-details { font-size: 24px; color: #4a5568; }
            .tv-days { font-size: 35px; font-weight: bold; float: left; margin-top: -10px; text-align: center; }
        </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 style='text-align: center; font-size: 60px; margin-bottom: 40px;'>شاشة متابعة الورشة (Live Queue)</h1>", unsafe_allow_html=True)

    if not ledger_df.empty:
        open_jobs = ledger_df[~ledger_df['case_status'].astype(str).eq(CASE_STATUS_CLOSED)]
        waiting_collection_df = ledger_df[ledger_df['collection_status'].isin([COLLECTION_AWAITING, COLLECTION_SPECIAL_AWAITING])] if not ledger_df.empty else pd.DataFrame()
        partner_claim_df = ledger_df[ledger_df['partner_claim_status'].eq('بانتظار مطالبة الشريك')] if not ledger_df.empty else pd.DataFrame()
        special_df = ledger_df[ledger_df['special_case'].astype(str).str.strip() != ''] if not ledger_df.empty else pd.DataFrame()

        display_items = []
        for _, r in open_jobs.iterrows():
            try:
                logged_dt = pd.to_datetime(str(r['date_logged']).split(' ')[0])
                if logged_dt > datetime.now(): logged_dt = datetime.now() - pd.Timedelta(days=2)
                days = (datetime.now() - logged_dt).days
                if days < 0: days = 0
            except:
                days = 0

            is_urgent = "عاجل" in str(r.get('priority', ''))

            display_items.append({
                "days": days,
                "urgent": is_urgent,
                "sid": r['service_id'],
                "tool": r['tool_name'],
                "issue": r['reported_issue'],
                "status": r['status'],
                "remarks": r.get('remarks', ''),
                "collection_status": r.get('collection_status', ''),
                "partner_claim_status": r.get('partner_claim_status', ''),
                "special_case": r.get('special_case', '')
            })

        display_items = sorted(display_items, key=lambda x: (not x['urgent'], -x['days']))

        if display_items:
            for item in display_items:
                if item['urgent']:
                    card_class = "tv-card-urgent"
                    tag = "🔥 عاجل جداً"
                    color = "#e53e3e"
                elif item['days'] >= 3:
                    card_class = "tv-card-delayed"
                    tag = "⚠️ متأخر"
                    color = "#dd6b20"
                else:
                    card_class = "tv-card-normal"
                    tag = "⚙️ قيد العمل"
                    color = "#3182ce"

                st.markdown(f"""
                <div class="{card_class}">
                    <div class="tv-days" style="color: {color};">{item['days']}<br><span style="font-size:16px;">أيام</span></div>
                    <div class="tv-title">{tag} | {item['sid']} - {item['tool']}</div>
                    <div class="tv-details"><b>العطل:</b> {item['issue']} <br> <b>الحالة:</b> {item['status']} <br> <b>الاستلام:</b> {item['collection_status']} <br> <b>ملاحظة خاصة:</b> {item['special_case']} <br> <b>مطالبة الشريك:</b> {item['partner_claim_status']} <br> <b>ملاحظات:</b> {item['remarks']}</div>
                </div>
                """, unsafe_allow_html=True)

                # Interactive Workshop Quick-Update Form
                with st.expander(f"⚡ تحديث سريع للسند ({item['sid']})"):
                    with st.form(f"quick_form_{item['sid']}"):
                        q_action = st.selectbox("الإجراء:", ["تحديث الملاحظات فقط", "إنجاز وجاهز للتسليم (Ready)"], key=f"act_{item['sid']}")
                        q_remark = st.text_input("إضافة ملاحظة ورشة:", value=item['remarks'], key=f"rem_{item['sid']}")

                        if st.form_submit_button("حفظ التحديث (Save)", use_container_width=True):
                            idx = ledger_df.index[ledger_df['service_id'] == item['sid']][0]
                            ledger_df.at[idx, 'remarks'] = q_remark
                            if "إنجاز" in q_action:
                                ledger_df.at[idx, 'document_origin'] = "مبيع خ ص: (جاهز ومفوتر)"
                                ledger_df.at[idx, 'status'] = "جاهز للتسليم (Ready)"
                                ledger_df.at[idx, 'date_resolved'] = datetime.now().strftime("%Y-%m-%d")
                            save_doctype("Ledger", ledger_df)
                            st.success("✅ تم التحديث بنجاح!")
                            st.rerun()

        st.divider()
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("📦 بانتظار تأكيد الاستلام", len(waiting_collection_df))
        with c2: st.metric("💼 مطالبات شركاء معلقة", len(partner_claim_df))
        with c3: st.metric("⚠️ حالات خاصة", len(special_df))
        if not partner_claim_df.empty:
            st.warning("💼 توجد حالات V لم تُسجّل لها خ صيانة بعد — راجع مطالبة الشريك.")
        if not special_df.empty:
            st.info("ℹ️ الحالات الخاصة مثل كفالة/رفض/مكلف/يعمل من المصدر لا تنتظر قبضاً؛ حالات العملاء تنتظر تأكيد الاستلام من الشاشة، وحالات الشريك تغلق تلقائياً.")
        else:
            st.markdown("<h1 style='text-align: center; color: #38a169; margin-top: 100px;'>✅ لا توجد أجهزة قيد الصيانة. الورشة خالية!</h1>", unsafe_allow_html=True)
    else:
        st.markdown("<h1 style='text-align: center; color: #38a169; margin-top: 100px;'>✅ العمل ممتاز! لا توجد مهام حالياً.</h1>", unsafe_allow_html=True)

# ==========================================
# MODULE 3: SUPPORT & MAINTENANCE
# ==========================================
elif st.session_state['current_module'] == 'Support':
    st.title("🛠️ وحدة الدعم والصيانة (Support Desk)")
    tab1, tab2, tab3 = st.tabs(["➕ بطاقة صيانة جديدة (New Ticket)", "🔄 تحديث الملف (Update & View)", "⚠️ قائمة المهام (SLA / Queue)"])

    with tab1:
        st.markdown("<div class='erp-card'>", unsafe_allow_html=True)
        with st.form("intake_form", clear_on_submit=True):
            st.subheader("تفاصيل استلام جهاز (Intake Form)")

            c_amn1, c_amn2, c_amn3 = st.columns(3)
            with c_amn1:
                warranty = st.selectbox("حالة الكفالة (Warranty)", ["خارج الكفالة", "ضمن كفالة"])
            with c_amn2:
                priority = st.selectbox("أولوية العمل (Priority)", ["عادي (Normal)", "عاجل 🔥 (Rush Job)"])
            with c_amn3:
                loaner = st.text_input("جهاز بديل معار للزبون (Loaner Item S/N - اختياري)")

            st.divider()

            c1, c2, c3 = st.columns(3)
            with c1:
                branch_select = st.selectbox("الفرع (Branch Prefix)", ["صيدا (S)", "درعا (D)", "وكيل / شريك (V)"])
                c_name = st.text_input("اسم الزبون (Customer Name)")
            with c2:
                c_phone = st.text_input("رقم الهاتف (Phone)")
                t_name_dropdown = st.selectbox("الجهاز (Item Lookup)", options=["أخرى (إدخال يدوي)"] + stock_list)
                t_name_manual = st.text_input("اسم الجهاز اليدوي (Manual Entry)")
            with c3:
                doc_origin = st.selectbox("الحالة المحاسبية (Origin)", ["اد خ ص: (استلام للصيانة)", "مبيع خ ص: (جاهز ومفوتر)", "خ صيانة: (تحميل على الوكيل)"])
                accessories = st.text_input("الملحقات المستلمة (Accessories) ⚠️ إلزامي", placeholder="مثال: بطارية، شاحن، حقيبة أو 'لا يوجد'")

            issue = st.text_area("العطل المرصود (Reported Issue)")

            st.markdown("📷 **التوثيق البصري (Media Documentation)**")
            photo_buffer = st.camera_input("التقاط صورة للجهاز أو الملحقات كإثبات حالة (Take Photo)")

            if st.form_submit_button("إنشاء السند (Create Document)", use_container_width=True):
                final_t_name = t_name_manual if t_name_dropdown == "أخرى (إدخال يدوي)" and t_name_manual else t_name_dropdown

                if not accessories.strip():
                    st.error("❌ حقل 'الملحقات المستلمة' إلزامي لمنع فقدان الأغراض. (اكتب 'لا يوجد' إن لم يسلمك شيء).")
                elif c_name and final_t_name and final_t_name != "أخرى (إدخال يدوي)":
                    branch_code = "S" if "S" in branch_select else "D" if "D" in branch_select else "V"
                    auto_id = generate_next_id(branch_code, ledger_df)
                    date_now = datetime.now().strftime("%Y-%m-%d")

                    photo_url = upload_to_cloud(photo_buffer) if photo_buffer else ""

                    new_row = {
                        "service_id": auto_id, "tool_name": final_t_name, "customer_name": c_name, "phone_number": c_phone,
                        "warranty_status": warranty, "document_origin": doc_origin, "reported_issue": issue,
                        "technician": current_user, "status": map_document_to_status(doc_origin, 0.0), "cost_debit": 0.0, "payment_credit": 0.0,
                        "balance": 0.0, "spare_parts": "لا حاجة / متوفرة", "resolution_notes": "", "remarks": "",
                        "date_logged": date_now, "date_resolved": "",
                        "accessories": accessories, "loaner_item": loaner, "priority": priority, "tool_photo_link": photo_url,
                        "source_account": "", "source_account_g": "", "source_document_count": 1, "document_history": doc_origin,
                        "repair_stage": "قيد المعالجة", "collection_status": COLLECTION_NOT_READY, "special_case": "", "partner_claim_status": "غير منطبق", "partner_claim_amount": 0.0,
                        "case_status": CASE_STATUS_OPEN, "closed_at": "", "closed_by": "", "close_note": ""
                    }
                    save_doctype("Ledger", pd.concat([ledger_df, pd.DataFrame([new_row])], ignore_index=True))
                    st.success(f"✅ تم إنشاء السند بنجاح برقم: {auto_id}")
                else:
                    st.error("❌ يرجى تعبئة اسم الزبون واسم الجهاز.")
        st.markdown("</div>", unsafe_allow_html=True)

    with tab2:
        if not ledger_df.empty:
            opts = {f"{r.get('priority', '')} {r['service_id']} - {r['customer_name']}": r['service_id'] for _, r in ledger_df.iterrows()}
            sel_id = opts[st.selectbox("ابحث عن السند (Search Document):", list(opts.keys()))]
            row_data = ledger_df[ledger_df['service_id'] == sel_id].iloc[0]

            is_locked = str(row_data.get('case_status', CASE_STATUS_OPEN)) == CASE_STATUS_CLOSED
            is_warranty = "ضمن" in str(row_data.get('warranty_status', ''))
            is_partner_ticket = str(sel_id).upper().startswith("V")

            if is_locked:
                st.markdown("<div class='locked-card'>", unsafe_allow_html=True)
                st.markdown(f"### 🔒 مستند مغلق (Submitted/Locked)")
                st.write(f"**رقم السند:** {sel_id} | **الزبون:** {row_data['customer_name']}")
                st.write(f"**حالة الملف:** {row_data['status']}")
                st.write(f"**أغلق بواسطة:** {row_data.get('closed_by', 'غير مسجل')} في {row_data.get('closed_at', row_data.get('date_resolved', ''))}")
                if row_data.get('close_note'): st.write(f"**ملاحظة الإغلاق:** {row_data.get('close_note')}")
                st.write(f"**التكلفة النهائية:** ${float(row_data['cost_debit']):.2f}")
                if row_data.get('tool_photo_link'): st.markdown(f"[📸 عرض صورة الجهاز عند الاستلام]({row_data['tool_photo_link']})")
                st.write("هذا الملف مغلق نهائياً لحماية القيود المالية. للطباعة يرجى التوجه لقسم المحاسبة.")
                st.markdown("</div>", unsafe_allow_html=True)
            else:
                st.markdown("<div class='erp-card'>", unsafe_allow_html=True)

                c_info1, c_info2 = st.columns(2)
                with c_info1:
                    st.caption("الملحقات المستلمة (Accessories):")
                    st.write(f"🎒 {row_data.get('accessories', 'غير مسجل')}")
                with c_info2:
                    if row_data.get('tool_photo_link'): st.markdown(f"📸 [عرض الصورة المرفقة للصيانة]({row_data['tool_photo_link']})")

                with st.form("update_form"):
                    doc_options = ["اد خ ص: (استلام للصيانة)", "مبيع خ ص: (جاهز ومفوتر)", "قبض د: (مدفوع ومسلم)", "قبض م: (مدفوع ومسلم)", "خ صيانة: (تحميل على الوكيل)"]
                    try: curr_i = [i for i, o in enumerate(doc_options) if str(row_data['document_origin']) in o][0]
                    except: curr_i = 0

                    c_a, c_b = st.columns(2)
                    with c_a: new_doc = st.selectbox("تحديث الحالة المحاسبية:", doc_options, index=curr_i)
                    with c_b:
                        sp_opts = ["لا حاجة / متوفرة", "بانتظار شحن مجاني", "بانتظار شحن عادي"]
                        try: sp_i = sp_opts.index(str(row_data['spare_parts']))
                        except: sp_i = 0
                        new_spare = st.selectbox("حالة قطع الغيار:", sp_opts, index=sp_i)

                    if is_warranty:
                        st.info("🛡️ هذا الجهاز ضمن الكفالة، تم تصفير التكلفة تلقائياً.")
                        cost = 0.0
                        pay = 0.0
                    else:
                        col1, col2, col3 = st.columns(3)
                        with col1: cost = st.number_input("التكلفة (Debit)", value=float(row_data['cost_debit'] or 0), step=1.0)
                        with col2: pay = st.number_input("الدفعة (Credit)", value=float(row_data['payment_credit'] or 0), step=1.0)
                        with col3: st.metric("الرصيد المتبقي (Balance)", f"${cost - pay:.2f}")

                    current_collection = str(row_data.get("collection_status", ""))
                    st.info(f"📄 حالة كشف الأمين: {current_collection or COLLECTION_NOT_READY}")
                    special_options = ["", "يعمل من المصدر", "الزبون رفض الإصلاح", "مكلف", "كفالة", "غير قابل للإصلاح", "لا يوجد عطل"]
                    current_special = str(row_data.get("special_case", ""))
                    try: special_i = special_options.index(current_special)
                    except: special_i = 0
                    new_special = st.selectbox("حالة خاصة (Special Case):", special_options, index=special_i)
                    partner_options = ["غير منطبق", "بانتظار مطالبة الشريك", "تمت مطالبة الشريك / تم التحصيل", "لا مطالبة - حالة خاصة"]
                    current_partner = str(row_data.get("partner_claim_status", "غير منطبق"))
                    try: partner_i = partner_options.index(current_partner)
                    except: partner_i = 0
                    new_partner_claim = st.selectbox("حالة مطالبة الشريك (Partner Claim):", partner_options, index=partner_i)

                    if "قبض" in new_doc:
                        new_collection = COLLECTION_CLOSED
                    elif new_special:
                        new_collection = COLLECTION_SPECIAL_AWAITING
                    elif "مبيع خ ص" in new_doc:
                        new_collection = COLLECTION_AWAITING
                    else:
                        new_collection = current_collection or COLLECTION_NOT_READY

                    new_status = map_document_to_status(new_doc, cost, new_collection, new_special, new_partner_claim, CASE_STATUS_OPEN)
                    c_n1, c_n2 = st.columns(2)
                    with c_n1: notes = st.text_area("ملاحظات الإصلاح (Resolution)", value=str(row_data.get('resolution_notes', '')))
                    with c_n2: remarks_update = st.text_area("تحديثات إضافية (Remarks)", value=str(row_data.get('remarks', '')))

                    if st.form_submit_button("تحديث السجل (Update Document)", use_container_width=True):
                        idx = ledger_df.index[ledger_df['service_id'] == sel_id][0]
                        ledger_df.at[idx, 'cost_debit'] = cost
                        ledger_df.at[idx, 'payment_credit'] = pay
                        ledger_df.at[idx, 'balance'] = cost - pay
                        ledger_df.at[idx, 'resolution_notes'] = notes
                        ledger_df.at[idx, 'remarks'] = remarks_update
                        ledger_df.at[idx, 'document_origin'] = new_doc
                        ledger_df.at[idx, 'status'] = new_status
                        ledger_df.at[idx, 'spare_parts'] = new_spare
                        ledger_df.at[idx, 'collection_status'] = new_collection
                        ledger_df.at[idx, 'special_case'] = new_special
                        ledger_df.at[idx, 'partner_claim_status'] = new_partner_claim
                        ledger_df.at[idx, 'repair_stage'] = 'جاهز للتسليم' if 'مبيع خ ص' in new_doc or 'قبض' in new_doc or new_special else 'قيد المعالجة'
                        save_doctype("Ledger", ledger_df)
                        st.success("✅ تم التحديث بنجاح!")
                        st.rerun()

                # Normal S/D cases are closed by قبض in Ameen. Ameen zero-value
                # sales invoices now close automatically as well. Manual closure is
                # reserved for special/no-charge customer cases where neither signal exists.
                special_collection_eligible = (
                    not is_partner_ticket
                    and (bool(normalize_doc_string(row_data.get('special_case', ''))) or is_warranty)
                )
                if special_collection_eligible:
                    st.markdown("### ✅ تأكيد استلام الحالة الخاصة")
                    st.caption("هذه الحالة لا تنتظر قبضاً في الأمين. أغلقها هنا فقط بعد أن يستلم الزبون الجهاز فعلياً.")
                    with st.form("close_case_form"):
                        close_note = st.text_input("ملاحظة الإغلاق (اختياري)")
                        if st.form_submit_button("تم الاستلام — إغلاق الحالة", use_container_width=True):
                            idx = ledger_df.index[ledger_df['service_id'] == sel_id][0]
                            closed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            ledger_df.at[idx, 'case_status'] = CASE_STATUS_CLOSED
                            ledger_df.at[idx, 'closed_at'] = closed_at
                            ledger_df.at[idx, 'closed_by'] = current_user
                            ledger_df.at[idx, 'close_note'] = close_note or f"استلام يدوي - {row_data.get('special_case', 'كفالة')}"
                            ledger_df.at[idx, 'collection_status'] = COLLECTION_CLOSED
                            ledger_df.at[idx, 'repair_stage'] = CASE_STATUS_CLOSED
                            ledger_df.at[idx, 'status'] = "مغلق - تم الاستلام (Closed)"
                            ledger_df.at[idx, 'date_resolved'] = closed_at.split(" ")[0]
                            save_doctype("Ledger", ledger_df)
                            st.success("✅ تم إغلاق الحالة نهائياً بعد تأكيد الاستلام.")
                            st.rerun()
                elif not is_partner_ticket:
                    st.info("ℹ️ الحالات العادية تغلق تلقائياً عند ظهور قبض في كشف الأمين. كما أن فاتورة مبيع خ ص بقيمة صفر تغلق تلقائياً حسب القاعدة الحالية.")

                if row_data['phone_number']:
                    phone_clean = re.sub(r'\D', '', str(row_data['phone_number']))
                    wa_msg = f"مرحباً {row_data['customer_name']}, جهازك ({row_data['tool_name']}) جاهز للاستلام من القصر الذهبي."
                    wa_link = f"https://wa.me/{phone_clean}?text={urllib.parse.quote(wa_msg)}"
                    st.markdown(f"<a href='{wa_link}' target='_blank'><button style='background-color:#25D366; color:white; border:none; padding:10px 20px; border-radius:5px; cursor:pointer; width:100%; font-size:16px; font-weight:bold;'>💬 إرسال إشعار للزبون (WhatsApp)</button></a>", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)
        else: st.info("لا توجد ملفات حالياً. يرجى إنشاء سند صيانة جديد.")

    with tab3:
        st.markdown("<div class='erp-card'>", unsafe_allow_html=True)
        if not ledger_df.empty:
            open_jobs = ledger_df[~ledger_df['case_status'].astype(str).eq(CASE_STATUS_CLOSED)]
            alerts = []
            for _, r in open_jobs.iterrows():
                try:
                    logged_dt = pd.to_datetime(str(r['date_logged']).split(' ')[0])
                    if logged_dt > datetime.now(): logged_dt = datetime.now() - pd.Timedelta(days=2)
                    days = (datetime.now() - logged_dt).days
                    if days < 0: days = 0
                except:
                    days = 0

                is_urgent = "عاجل" in str(r.get('priority', ''))

                alert = "✅ طبيعي"
                if "المعالجة" in str(r['status']) or "الانتظار" in str(r['status']):
                    if days > 5: alert = "🔴 متأخر جداً"
                    elif days > 3: alert = "🟠 متأخر"
                elif "جاهز" in str(r['status']) and days > 7: alert = "🔴 تأخر بالاستلام"

                alerts.append({
                    "أولوية": "عاجل 🔥" if is_urgent else "عادي",
                    "الحالة (SLA)": alert,
                    "أيام التوقف": days,
                    "السند": r['service_id'],
                    "الزبون": r['customer_name'],
                    "الجهاز": r['tool_name'],
                    "الوضع": r['status']
                })

            df_alerts = pd.DataFrame(alerts)
            if not df_alerts.empty:
                c_filt1, c_filt2 = st.columns([3, 1])
                with c_filt1:
                    all_statuses = df_alerts['الوضع'].unique().tolist()
                    selected_statuses = st.multiselect("🔍 تصفية حسب الوضع (Filter by Status):", options=all_statuses, default=all_statuses)

                filtered_alerts = df_alerts[df_alerts['الوضع'].isin(selected_statuses)]

                with c_filt2:
                    st.metric("العدد (Count)", len(filtered_alerts))

                filtered_alerts = filtered_alerts.sort_values(by=["أولوية", "أيام التوقف"], ascending=[False, False])
                st.dataframe(filtered_alerts, use_container_width=True)
            else:
                st.success("✅ ممتاز! جميع الأجهزة جاهزة أو تم تسليمها، ولا توجد مهام متأخرة أو قيد المعالجة.")
        else:
            st.info("📂 قاعدة البيانات فارغة. يرجى إنشاء سند جديد أو استيراد ملف الأمين.")
        st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# MODULE 4: STOCK & INVENTORY
# ==========================================
elif st.session_state['current_module'] == 'Stock':
    st.title("📦 وحدة المستودعات والمخزون (Stock Module)")
    st.markdown("<div class='erp-card'>", unsafe_allow_html=True)
    if not stock_df.empty:
        stock_df['quantity'] = pd.to_numeric(stock_df['quantity'], errors='coerce').fillna(0)
        low_stock = stock_df[stock_df['quantity'] <= 2]
        if not low_stock.empty:
            st.error(f"⚠️ يوجد {len(low_stock)} أصناف تتطلب إعادة طلب (Reorder Alert).")

    c1, c2 = st.columns(2)
    with c1: st.download_button("📥 تصدير السجل (Export)", data=convert_df_to_excel(stock_df) if not stock_df.empty else b"", file_name="Stock_Master.xlsx", use_container_width=True)
    with c2:
        with st.expander("📤 استيراد لائحة الأسعار (Import List)"):
            uploaded_stock = st.file_uploader("رفع ملف Excel", type=["xlsx"])
            if uploaded_stock and st.button("استيراد (Import)"):
                raw = pd.read_excel(uploaded_stock)
                if 'MtCode' in raw.columns and 'اسم المادة' in raw.columns:
                    p_col = 'الجملة' if 'الجملة' in raw.columns else raw.columns[-1]
                    new_items = raw[['MtCode', 'اسم المادة', p_col]].copy()
                    new_items.columns = ['item_code', 'item_name', 'price']
                    new_items = new_items.dropna(subset=['item_code'])
                    new_items['price'] = pd.to_numeric(new_items['price'], errors='coerce').fillna(0.0)
                    new_items['quantity'] = 0
                    if not stock_df.empty:
                        q_dict = dict(zip(stock_df['item_code'], stock_df['quantity']))
                        new_items['quantity'] = new_items['item_code'].map(q_dict).fillna(0)
                    save_doctype("Stock", new_items[STOCK_COLUMNS])
                    st.success("✅ اكتمل الاستيراد.")
                    st.rerun()

    if not stock_df.empty:
        edited_stock = st.data_editor(stock_df, num_rows="dynamic", use_container_width=True)
        if st.button("💾 حفظ التعديلات (Save Stock)", use_container_width=True):
            save_doctype("Stock", edited_stock)
            st.success("تم الحفظ!")
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# MODULE 5: LOGISTICS (HAWARA & DISPATCH)
# ==========================================
elif st.session_state['current_module'] == 'Logistics':
    st.title("🚚 وحدة الشحن واللوجستيات (Logistics)")
    tab1, tab2 = st.tabs(["📑 مشتريات وشحن حوارة (Supplier Orders)", "📦 التوصيل المحلي (Local Dispatch)"])

    with tab1:
        st.markdown("<div class='erp-card'>", unsafe_allow_html=True)
        with st.form("hawara_form", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                h_id = st.text_input("معرف الطلب (Order ID)")
                h_type = st.selectbox("نوع العملية (Type)", ["طلب قطع غيار", "إرسال للصيانة", "استرجاع بضاعة"])
                linked_sid = st.selectbox("ارتباط بسند (Link to Ticket)", options=["بدون ربط"] + ledger_df['service_id'].tolist() if not ledger_df.empty else ["بدون ربط"])
            with c2:
                courier = st.selectbox("شركة الشحن (Courier)", ["شركة أرامكس", "نقل قدموس", "ساعي داخلي", "شركة حوارة"])
                h_note = st.text_input("بوليصة الشحن (Delivery Note)")
            with c3:
                h_status = st.selectbox("الحالة (Status)", ["قيد الطلب", "في الطريق", "تم الاستلام"])
                uploaded_doc = st.file_uploader("مرفق الفاتورة (Invoice Image)", type=["png", "jpg", "jpeg"])

            if st.form_submit_button("حفظ الطلبية (Submit Order)", use_container_width=True):
                if h_id:
                    file_url = upload_to_cloud(uploaded_doc) if uploaded_doc else ""
                    new_hawara = {
                        "order_id": h_id, "order_type": h_type, "linked_service_id": linked_sid if linked_sid != "بدون ربط" else "",
                        "courier": courier, "delivery_note": h_note, "document_link": file_url, "status": h_status, "date_logged": datetime.now().strftime("%Y-%m-%d")
                    }
                    save_doctype("Hawara", pd.concat([hawara_df, pd.DataFrame([new_hawara])], ignore_index=True))
                    st.success("✅ تم حفظ طلبية حوارة بنجاح!")
                else: st.error("يرجى إدخال معرف الطلب.")

        if not hawara_df.empty:
            st.subheader("سجل الطلبيات (Order Log)")
            edited_hawara = st.data_editor(hawara_df, num_rows="dynamic", use_container_width=True, column_config={"document_link": st.column_config.LinkColumn("المرفق", display_text="🔗 عرض")})
            if st.button("حفظ التعديلات (Save Edits)"): save_doctype("Hawara", edited_hawara); st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    with tab2:
        st.markdown("<div class='erp-card'>", unsafe_allow_html=True)
        with st.form("dispatch_form", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                disp_id = st.text_input("رقم الإرسالية (Dispatch ID)")
                ready_list = ledger_df[ledger_df['status'].str.contains('جاهز', na=False) & ~ledger_df['case_status'].astype(str).eq(CASE_STATUS_CLOSED)] if not ledger_df.empty else pd.DataFrame()
                sel_service = st.selectbox("الجهاز (Ready Tool)", options=ready_list['service_id'].tolist() if not ready_list.empty else [])
            with c2:
                disp_courier = st.selectbox("شركة النقل (Courier)", ["شركة أرامكس", "نقل قدموس", "ساعي داخلي"])
                disp_note = st.text_input("رقم البوليصة (Tracking No)")
            with c3:
                disp_file = st.file_uploader("مرفق البوليصة (Receipt Image)", type=["png", "jpg", "jpeg"])

            if st.form_submit_button("حفظ الإرسالية (Submit Dispatch)", use_container_width=True):
                if disp_id and sel_service:
                    file_url = upload_to_cloud(disp_file) if disp_file else ""
                    cust = ledger_df[ledger_df['service_id'] == sel_service].iloc[0]['customer_name'] if not ledger_df.empty else ''
                    new_disp = {
                        "dispatch_id": disp_id, "service_id": sel_service, "customer_name": cust,
                        "courier": disp_courier, "delivery_note": disp_note, "document_link": file_url, "date": datetime.now().strftime("%Y-%m-%d")
                    }
                    save_doctype("Dispatch", pd.concat([dispatch_df, pd.DataFrame([new_disp])], ignore_index=True))
                    st.success("✅ تم تسجيل الإرسالية وحفظ المرفق بنجاح!")
                else: st.error("يرجى إدخال رقم الإرسالية وسند الصيانة.")

        if not dispatch_df.empty:
            st.data_editor(dispatch_df, num_rows="dynamic", use_container_width=True, column_config={"document_link": st.column_config.LinkColumn("الإيصال", display_text="🔗 عرض")})
        st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# MODULE 6: ACCOUNTING & INVOICING
# ==========================================
elif st.session_state['current_module'] == 'Accounting':
    st.title("💰 الإدارة المالية والمحاسبة (Accounting)")

    tab1, tab2, tab3 = st.tabs(["🧾 طباعة الفواتير (Print Invoices)", "📊 التقارير ودفتر الأستاذ (General Ledger)", "⚙️ استيراد البيانات (Data Import)"])

    with tab1:
        st.markdown("<div class='erp-card'>", unsafe_allow_html=True)
        if not ledger_df.empty:
            inv_opts = ledger_df['service_id'].unique().tolist()
            sel_inv = st.selectbox("اختر رقم السند لتوليد الفاتورة (Select Ticket for Invoice):", options=inv_opts)

            if st.button("🖨️ توليد الفاتورة (Generate Invoice)", use_container_width=True):
                inv_data = ledger_df[ledger_df['service_id'] == sel_inv].iloc[0]

                invoice_html = f"""
                <div class="invoice-box">
                    <div class="invoice-header">
                        <h2>القصر الذهبي للمعدات الصناعية</h2>
                        <p>Al-Qasr Al-Zahabi | صيانة - بيع - تأجير</p>
                    </div>
                    <table style="width:100%; margin-bottom:20px; text-align:right; direction:rtl;">
                        <tr>
                            <td><b>رقم السند:</b> {inv_data['service_id']}</td>
                            <td><b>التاريخ:</b> {datetime.now().strftime("%Y-%m-%d")}</td>
                        </tr>
                        <tr>
                            <td><b>الزبون:</b> {inv_data['customer_name']}</td>
                            <td><b>الهاتف:</b> {inv_data['phone_number']}</td>
                        </tr>
                    </table>
                    <hr>
                    <table style="width:100%; text-align:right; direction:rtl; border-collapse: collapse; margin-top:20px;">
                        <tr style="background:#f7fafc; border-bottom:1px solid #cbd5e0;">
                            <th style="padding:10px;">البيان (Description)</th>
                            <th style="padding:10px;">المبلغ (Amount)</th>
                        </tr>
                        <tr style="border-bottom:1px solid #edf2f7;">
                            <td style="padding:10px;">صيانة أداة: {inv_data['tool_name']}<br><small>ملاحظات: {inv_data['resolution_notes']}</small></td>
                            <td style="padding:10px;">${float(inv_data['cost_debit']):.2f}</td>
                        </tr>
                        <tr style="border-bottom:1px solid #edf2f7;">
                            <td style="padding:10px;">الدفعة المقدمة (Credit)</td>
                            <td style="padding:10px;">${float(inv_data['payment_credit']):.2f}</td>
                        </tr>
                        <tr style="font-weight:bold; background:#ebf8ff;">
                            <td style="padding:10px;">الرصيد المتبقي (Total Due)</td>
                            <td style="padding:10px;">${float(inv_data['balance']):.2f}</td>
                        </tr>
                    </table>
                    <div style="text-align:center; margin-top:40px; font-size:12px; color:#718096;">
                        شكراً لتعاملكم معنا. (Thank you for your business.)<br>
                        <i>يمكن طباعة هذه الصفحة باستخدام (Ctrl + P)</i>
                    </div>
                </div>
                """
                st.components.v1.html(invoice_html, height=600, scrolling=True)
                st.info("💡 اضغط `Ctrl + P` أو `Cmd + P` في المتصفح لطباعة الفاتورة أو حفظها كـ PDF.")
        st.markdown("</div>", unsafe_allow_html=True)

    with tab2:
        st.markdown("<div class='erp-card'>", unsafe_allow_html=True)
        if not ledger_df.empty:
            view_type = st.radio("نوع العرض (View Type)", ["دفتر الأستاذ العام (General Ledger)", "حسابات الشركاء والوكلاء (Partner Ledger)"])

            if view_type == "دفتر الأستاذ العام (General Ledger)":
                st.dataframe(ledger_df, use_container_width=True)
                st.download_button("📥 تصدير الدفتر (Export Ledger)", data=convert_df_to_excel(ledger_df), file_name="General_Ledger.xlsx")
            else:
                partners_df = ledger_df[ledger_df['service_id'].str.startswith('V', na=False) | ledger_df['document_origin'].str.contains('خ صيانة', na=False)]
                st.dataframe(partners_df[['service_id', 'customer_name', 'tool_name', 'cost_debit', 'balance']], use_container_width=True)
                st.metric("مجموع ذمم الوكلاء", f"${partners_df['balance'].astype(float).sum():.2f}")
        st.markdown("</div>", unsafe_allow_html=True)

    with tab3:
        if is_admin:
            st.markdown("<div class='erp-card'>", unsafe_allow_html=True)
            st.subheader("📥 استيراد كشوفات الأمين (Legacy Import Tool)")
            uploaded_legacy = st.file_uploader("رفع ملف Excel", type=["xlsx"])
            if uploaded_legacy and st.button("تنفيذ الاستيراد (Run Import)"):
                with st.spinner("Processing Ameen repair ledger..."):
                    raw_excel = pd.read_excel(uploaded_legacy, sheet_name=0, header=0)
                    imported_df = normalize_ameen_dataframe(raw_excel)

                    if imported_df.empty:
                        st.error("❌ لم يتم العثور على سجلات صيانة صالحة في أعمدة A:L.")
                    else:
                        # Merge by service_id so repeated Ameen lines become one repair case.
                        existing = ledger_df.copy()
                        if not existing.empty:
                            existing = apply_workflow_columns(existing)
                            imported_df = preserve_manual_closures(imported_df, existing)
                            existing = existing[~existing["service_id"].astype(str).isin(imported_df["service_id"].astype(str))]
                        merged = pd.concat([existing, imported_df], ignore_index=True)
                        merged = apply_workflow_columns(merged)
                        save_doctype("Ledger", merged)

                        st.success(f"✅ تم استيراد {len(imported_df)} حالات صيانة من كشف الأمين.")
                        c1, c2, c3, c4, c5 = st.columns(5)
                        c1.metric("قيد المعالجة", int((imported_df["repair_stage"] == "قيد المعالجة").sum()))
                        c2.metric("جاهزة للتسليم", int((imported_df["repair_stage"] == "جاهز للتسليم").sum()))
                        c3.metric("بانتظار الاستلام", int(imported_df["collection_status"].isin([COLLECTION_AWAITING, COLLECTION_SPECIAL_AWAITING]).sum()))
                        c4.metric("مطالبات الشركاء", int((imported_df["partner_claim_status"] == "بانتظار مطالبة الشريك").sum()))
                        c5.metric("حالات مغلقة", int((imported_df["case_status"] == CASE_STATUS_CLOSED).sum()))
                        st.dataframe(imported_df[["service_id", "customer_name", "tool_name", "document_origin", "repair_stage", "collection_status", "case_status", "closed_at", "closed_by", "close_note", "special_case", "partner_claim_status", "remarks"]], use_container_width=True)
                        st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
