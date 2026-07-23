import os
import time
import sys
import psycopg2
import requests
import unicodedata
from datetime import datetime

# ==========================================
# CLOUD BACKEND ENGINE CONFIGURATION
# ==========================================
# Reads securely from GitHub environment secrets instead of exposing it in raw text
DATABASE_URL = os.environ.get("NEON_DB_URL")
EXCLUDED_CODES = ("40", "10", "20")

# If you use a free WhatsApp gateway provider, put your API webhook url target here
GATEWAY_API_URL = "https://yourgatewayprovider.com" 

if not DATABASE_URL:
    print("❌ Critical configuration error: NEON_DB_URL environment variable is missing.")
    sys.exit(1)

def clean_phone(raw_phone):
    if not raw_phone: return ""
    clean_raw = unicodedata.normalize('NFKC', str(raw_phone)).encode('ascii', 'ignore').decode('ascii')
    phone = clean_raw.strip().replace(" ", "").replace("-", "").replace("+", "").lstrip("0")
    return f"963{phone}" if phone.startswith('9') and len(phone) == 9 else (f"963{phone[1:]}" if phone.startswith('09') else phone)

def clean_txt(raw_text):
    if not raw_text: return ""
    return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip())

def send_via_cloud_gateway(phone, message):
    """
    Sends the message text layout directly over internet networks.
    This replaces clipboard-copying and mouse-clicking commands entirely!
    """
    print(f"📡 Forwarding message to phone: {phone}")
    print(f"💬 Text: {message}")
    
    # Optional connection payload example:
    # try:
    #     requests.post(GATEWAY_API_URL, json={"to": phone, "body": message}, timeout=10)
    # except Exception as e:
    #     print(f"⚠️ Gateway transmission hold: {e}")

def main():
    print("📡 Establishing secure network connection channel pipeline to Neon Server...")
    conn = psycopg2.connect(DATABASE_URL)
    cursor = conn.cursor()
    
    cursor.execute("SELECT MAX(id) FROM iclock_transaction;")
    res = cursor.fetchone()
    last_processed_id = res[0] if res and res[0] else 0
    print(f"✅ Connection successful! Monitoring starting from system entry ID: {last_processed_id}")
    
    start_time = time.time()
    
    while True:
        # Gracefully exit right before hitting the hard 6-hour execution limit (315 minutes = 5 hours 15 mins)
        if (time.time() - start_time) > (315 * 60):
            print("⏳ Reached cloud engine sequence threshold limits. Re-routing execution cycle cleanly...")
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
                punch_count = count_res[0] if count_res and count_res[0] else 1
                
                if punch_count % 2 != 0:
                    status_msg = f"مرحباً {name_clean}، تم تسجيل بصمة *الدخول* بنجاح عند الساعة {time_str}. أتمنى لك يوماً سعيداً! ✨"
                else:
                    status_msg = f"مرحباً {name_clean}، تم تسجيل بصمة *الخروج* بنجاح عند الساعة {time_str}. رافقتك السلامة! 🏡"
                
                send_via_cloud_gateway(phone_clean, status_msg)
                last_processed_id = t_id
                
            time.sleep(5) # Rest 5 seconds between database scans
            
        except psycopg2.DatabaseError as db_err:
            conn.rollback()
            time.sleep(10)
        except Exception as e:
            time.sleep(10)

if __name__ == "__main__":
    main()
