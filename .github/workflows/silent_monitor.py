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

async def main():
    print("🚀 Initializing virtual background browser instance inside GitHub Cloud...")
    session_dir = r"C:\Users\runneradmin\AppData\Local\Google\Chrome\User Data\WhatsAppCloudSession"
    
    # Ensure directory path structures are allocated cleanly inside the container environment
    if not os.path.exists(session_dir):
        os.makedirs(session_dir)
        
    try:
        # FIXED: Added critical args required to bypass administrative execution blocking on GitHub Virtual Windows Servers
        browser = await launch(
            headless=True,
            userDataDir=session_dir,
            args=[
                '--no-sandbox', 
                '--disable-setuid-sandbox', 
                '--disable-dev-shm-usage',
                '--disable-gpu',
                '--window-size=1280,800'
            ]
        )
        page = await browser.newPage()
        await page.setUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    except Exception as browser_err:
        print(f"❌ Failed to spin up virtual Chromium engine process: {browser_err}")
        # Create an emergency placeholder file so the GitHub Artifact step doesn't break the build pipeline run
        with open("whatsapp_cloud_login.png", "w") as f:
            f.write("Browser launch crashed.")
        sys.exit(1)
        
    print("🌐 Connecting to WhatsApp Web core framework...")
    try:
        await page.goto("https://whatsapp.com", {'waitUntil': 'networkidle2', 'timeout': 60000})
        print("⏳ Settling connection states...")
        await asyncio.sleep(20)
    except Exception as nav_err:
        print(f"⚠️ Initial network link synchronization timed out: {nav_err}")
    
    # Save the auth state token layout capture checkpoint directly to project root path layout safely
    await page.screenshot({'path': 'whatsapp_cloud_login.png'})
    print("📸 Auth snapshot saved as 'whatsapp_cloud_login.png' inside working cloud directory root workspace.")
    
    # Connect to database and look up latest transaction ID
    try:
        conn = psycopg2.connect(DATABASE_URL)
        cursor = conn.cursor()
        cursor.execute("SELECT MAX(id) FROM iclock_transaction;")
        res = cursor.fetchone()
        last_processed_id = res[0] if res and res[0] else 0
        print(f"📡 Real-time cloud background monitor active. Tracking updates from ID: {last_processed_id}")
    except Exception as db_init_err:
        print(f"❌ Cloud DB handshake failure: {db_init_err}")
        await browser.close()
        sys.exit(1)
        
    start_time = time.time()
    
    while True:
        # Exit smoothly at 4 hours and 40 minutes (280 minutes) to let cache configurations save state safely
        if (time.time() - start_time) > (280 * 60):
            print("⏳ Reached runtime block cycle threshold constraint limits. Cycling process logs safely...")
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
                
                print(f"✉ Dispatching cloud headless payload directly to phone route: {phone_clean}")
                target_url = f"https://whatsapp.com/send?phone={phone_clean}&text={status_msg}"
                
                await page.goto(target_url, {'waitUntil': 'networkidle2'})
                await asyncio.sleep(8)  
                
                # Triggers the hidden click send key event cleanly via background JS contexts inside RAM
                await page.evaluate("""() => {
                    const sendBtn = document.querySelector('span[data-icon="send"]') || document.querySelector('button[aria-label="Send"]');
                    if(sendBtn) sendBtn.click();
                }""")
                await asyncio.sleep(2)
                print(f"✔ Done. Confirmation message routed for Employee Code: {emp_code}")
                
                last_processed_id = t_id
                
            await asyncio.sleep(5)
            
        except psycopg2.DatabaseError as db_err:
            print(f"⚠️ DB transaction pipeline stalled, rolling back channel tree branches: {db_err}")
            conn.rollback()
            await asyncio.sleep(10)
        except Exception as loop_err:
            print(f"⏳ Background engine padding idle sync check: {loop_err}")
            await asyncio.sleep(10)
            
    cursor.close()
    conn.close()
    await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
