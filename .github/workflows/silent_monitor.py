import os
import time
import sys
import psycopg2
import unicodedata
import requests
from datetime import datetime

# ==========================================
# 0. CLOUD ENGINE ROUTING CONFIGURATION
# ==========================================
DATABASE_URL = os.environ.get("NEON_DB_URL")
EXCLUDED_CODES = ("NONE",)

# Points straight back to your local office PC network tunnel setup
OFFICE_PC_BRIDGE_URL = os.environ.get("OFFICE_PC_BRIDGE_URL")

if not DATABASE_URL or not OFFICE_PC_BRIDGE_URL:
    print("❌ Critical configuration error: Required cloud environment variables are missing.")
    sys.exit(1)

def clean_phone(raw_phone):
    if not raw_phone: return ""
    clean_raw = unicodedata.normalize('NFKC', str(raw_phone)).encode('ascii', 'ignore').decode('ascii')
    phone = clean_raw.strip().replace(" ", "").replace("-", "").replace("+", "").lstrip("0")
    return f"963{phone}" if phone.startswith('9') and len(phone) == 9 else (f"963{phone[1:]}" if phone.startswith('09') else phone)

def clean_txt(raw_text):
    if not raw_text: return ""
    return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip())

def dispatch_to_office_pc(phone, message):
    """Pipes data text rows back to your office PC background tunnel instantly."""
    try:
        payload = {"phone": phone, "message": message}
        # THIS CUSTOM HEADER COMPLETELY BYPASSES THE FREE NGROK INTERCEPTION WARNING PAGE:
        custom_headers = {
            "ngrok-skip-browser-warning": "69420"
        }
        response = requests.post(OFFICE_PC_BRIDGE_URL, json=payload, headers=custom_headers, timeout=15)
        if response.status_code == 200:
            print(f"🚀 Packet successfully routed to office PC for WhatsApp delivery to: {phone}")
            return True
        else:
            print(f"⚠️ Office PC bridge returned status code: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ Failed to reach office PC endpoint across network: {e}")
        return False

def main():
    print("📡 Connecting to Neon Cloud data streams...")
    conn = psycopg2.connect(DATABASE_URL)
    cursor = conn.cursor()
    
    cursor.execute("SELECT MAX(id) FROM iclock_transaction;")
    res = cursor.fetchone()
    last_processed_id = res[0] if res and res[0] is not None else 0
    print(f"✅ Connection successful! Cloud monitoring from transaction ID: {last_processed_id}")
    
    start_time = time.time()
    
    while True:
        # Gracefully restart container runner before hitting the 6-hour execution limit (320 minutes)
        if (time.time() - start_time) > (320 * 60):
            print("⏳ Reached cloud loop cycle expiration constraints. Refreshing engine...")
            break
            
        try:
            today_str = datetime.now().strftime('%Y-%m-%d')
            
            query = f"""
                SELECT t.id, e.emp_code, e.first_name, e.mobile, (t.punch_time AT TIME ZONE 'GMT-3')
                FROM iclock_transaction t
                JOIN personnel_employee e ON t.emp_id = e.id
                WHERE t.id > {last_processed_id} AND e.emp_code NOT IN ({",".join(f"'{c}'" for c in EXCLUDED_CODES)})
                ORDER BY t.id ASC;
            """
            cursor.execute(query)
            new_punches = cursor.fetchall()
            
            for punch in new_punches:
                t_id, emp_code, first_name, mobile, punch_time = punch
                name_clean = clean_txt(first_name)
                phone_clean = clean_phone(mobile)
                time_str = punch_time.strftime('%I:%M:%S %p')
                
                if not phone_clean or phone_clean == "963":
                    last_processed_id = t_id
                    continue
                
                count_query = f"""
                    SELECT COUNT(id) FROM iclock_transaction 
                    WHERE emp_id = (SELECT id FROM personnel_employee WHERE emp_code = %s)
                    AND (punch_time AT TIME ZONE 'GMT-3')::date = %s AND id <= %s;
                """
                cursor.execute(count_query, (emp_code, today_str, t_id))
                count_res = cursor.fetchone()
                punch_count = count_res[0] if count_res else 1
                
                if punch_count % 2 != 0:
                    status_msg = f"مرحباً {name_clean}، تم تسجيل بصمة *الدخول* بنجاح عند الساعة {time_str}. أتمنى لك يوماً سعيداً! ✨"
                else:
                    status_msg = f"مرحباً {name_clean}، تم تسجيل بصمة *الخروج* بنجاح عند الساعة {time_str}. رافقتك السلامة! 🏡"
                
                dispatch_to_office_pc(phone_clean, status_msg)
                last_processed_id = t_id
                
            time.sleep(5)
            
        except psycopg2.DatabaseError as db_err:
            conn.rollback()
            time.sleep(5)
        except Exception:
            time.sleep(5)

if __name__ == "__main__":
    main()
