import os
import time
import sys
import asyncio
import psycopg2
import unicodedata
from datetime import datetime
from pyppeteer import launch

# ==========================================
# CLOUD BACKEND ENGINE CONFIGURATION
# ==========================================
DATABASE_URL = os.environ.get("NEON_DB_URL")
EXCLUDED_CODES = ("40", "10", "20")

if not DATABASE_URL:
    print("Critical configuration error: NEON_DB_URL environment variable is missing.")
    sys.exit(1)

def clean_phone(raw_phone):
    if not raw_phone: return ""
    clean_raw = unicodedata.normalize('NFKC', str(raw_phone)).encode('ascii', 'ignore').decode('ascii')
    phone = clean_raw.strip().replace(" ", "").replace("-", "").replace("+", "").lstrip("0")
    return f"963{phone}" if phone.startswith('9') and len(phone) == 9 else (f"963{phone[1:]}" if phone.startswith('09') else phone)

def clean_txt(raw_text):
    if not raw_text: return ""
    return str(unicodedata.normalize('NFKC', str(raw_text)).replace('\u2066','').replace('\u2069','').strip())

async def main():
    print("Initializing virtual background browser instance inside GitHub Cloud...")
    
    # Store session data cleanly in the runner's AppData profile folder
    session_dir = r"C:\Users\runneradmin\AppData\Local\Google\Chrome\User Data\WhatsAppCloudSession"
    
    browser = await launch(
        headless=True,
        args=['--no-sandbox', '--disable-setuid-sandbox', '--window-size=1280,800'],
        userDataDir=session_dir
    )
    page = await browser.newPage()
    await page.setUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    print("Connecting to WhatsApp Web core framework...")
    await page.goto("https://whatsapp.com")
    
    # Give the cloud browser plenty of time to boot and check session cookies
    await asyncio.sleep(25)
    
    # Save a verification image file to capture either your active chat panel or the QR code
    await page.screenshot({'path': 'whatsapp_cloud_login.png'})
    print("Auth snapshot saved as 'whatsapp_cloud_login.png'. Uploading it to workspace artifacts...")
    
    # Connect to database and look up latest transaction ID
    conn = psycopg2.connect(DATABASE_URL)
    cursor = conn.cursor()
    cursor.execute("SELECT MAX(id) FROM iclock_transaction;")
    res = cursor.fetchone()
    last_processed_id = res[0] if res and res[0] else 0
    print(f"Real-time background monitor active. Tracking updates from ID: {last_processed_id}")
    
    start_time = time.time()
    
    while True:
        # Exit smoothly at 4 hours and 40 minutes (280 minutes) to give GitHub time to encrypt and cache cookies
        if (time.time() - start_time) > (280 * 60):
            print("Cyclic duration threshold reached. Exporting state cache and restarting...")
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
                
                print(f"Sending headless cloud message payload to: {phone_clean}")
                target_url = f"https://whatsapp.com/send?phone={phone_clean}&text={status_msg}"
                
                await page.goto(target_url)
                await asyncio.sleep(8)  
                
                # Clicks the send button using virtual Javascript actions inside RAM
                await page.evaluate("""() => {
                    const sendBtn = document.querySelector('span[data-icon="send"]') || document.querySelector('button[aria-label="Send"]');
                    if(sendBtn) sendBtn.click();
                }""")
                await asyncio.sleep(2)
                print(f"✔ Done. Confirmation message routed for Employee Code: {emp_code}")
                
                last_processed_id = t_id
                
            await asyncio.sleep(5)
            
        except psycopg2.DatabaseError as db_err:
            conn.rollback()
            await asyncio.sleep(10)
        except Exception as e:
            await asyncio.sleep(10)
            
    cursor.close()
    conn.close()

if __name__ == "__main__":
    asyncio.run(main())
