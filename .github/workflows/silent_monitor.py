import os
import time
import sys
import asyncio
import psycopg2
import unicodedata
import urllib.parse
from datetime import datetime
from pyppeteer import launch

# ==========================================
# 0. ALIGNMENT STRINGS CONFIGURATION (UNICODE ESCAPED)
# ==========================================
MSG_TEMPLATES = {
    "db_err": "Database error encountered, rolling back: {}",
    "sys_err": "Sync placeholder pause: {}",
    "in_punch": "\u0645\u0631\u062d\u0628\u0627\u064b {}\u060c \u062a\u0645 \u062a\u0633\u062c\u064a\u0644 \u0628\u0635\u0645\u0629 *\u0627\u0644\u062f\u062e\u0648\u0644* \u0628\u0646\u062c\u0627\u062d \u0639\u0646\u062f \u0627\u0644\u0633\u0627\u0639\u0645 {} \u0622\u062a\u0645\u0646\u0649 \u0644\u0643 \u064a\u0641 \u064a\u0648\u0645\u0627\u064b \u0633\u0639\u064a\u062f\u0627\u064b! \u2728",
    "out_punch": "\u0645\u0631\u062d\u0628\u0627\u064b {}\u060c \u062a\u0645 \u062a\u0633\u062c\u064a\u0644 \u0628\u0635\u0645\u0629 *\u0627\u0644\u062e\u0631\u0648\u062c* \u0628\u0646\u062c\u0627\u062d \u0639\u0646\u062f \u0627\u0644\u0633\u0627\u0639\u0629 {} \u0631\u0627\u0641\u0642\u062a\u0643 \u0627\u0644\u0633\u0644\u0627\u0645\u0629! \ud83c\udfe1"
}

DATABASE_URL = os.environ.get("NEON_DB_URL")
EXCLUDED_CODES = ("40", "10", "20")

if not DATABASE_URL:
    print("Critical configuration error: NEON_DB_URL is missing.")
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
    print("Initializing virtual browser instance inside RAM...")
    session_dir = r"C:\Users\runneradmin\AppData\Local\Google\Chrome\User Data\WhatsAppCloudSession"
    
    if not os.path.exists(session_dir):
        os.makedirs(session_dir)
        
    try:
        # FIXED: Forces pyppeteer to launch using the native Microsoft Edge pre-installed on the GitHub Windows Runner
        browser = await launch(
            headless=True,
            userDataDir=session_dir,
            executablePath=r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
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
        print(f"Browser engine crash: {browser_err}")
        with open("whatsapp_cloud_login.png", "w") as f:
            f.write("Crash placeholder")
        sys.exit(1)
        
    print("Connecting to WhatsApp Web...")
    try:
        await page.goto("https://whatsapp.com", {'waitUntil': 'networkidle2', 'timeout': 60000})
        await asyncio.sleep(25)
    except Exception as nav_err:
        print(f"Navigation timeout: {nav_err}")
    
    await page.screenshot({'path': 'whatsapp_cloud_login.png'})
    print("Snapshot created successfully.")
    
    try:
        conn = psycopg2.connect(DATABASE_URL)
        cursor = conn.cursor()
        cursor.execute("SELECT MAX(id) FROM iclock_transaction;")
        res = cursor.fetchone()
        last_processed_id = res if res and res else 0
        print(f"Connected! Monitoring updates from ID: {last_processed_id}")
    except Exception as db_init_err:
        print(f"Database handshake failure: {db_init_err}")
        await browser.close()
        sys.exit(1)
        
    start_time = time.time()
    
    while True:
        if (time.time() - start_time) > (280 * 60):
            print("Cycling loop execution cleanly...")
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
                punch_count = count_res if count_res else 1
                
                if punch_count % 2 != 0:
                    status_msg = MSG_TEMPLATES["in_punch"].format(name_clean, time_str)
                else:
                    status_msg = MSG_TEMPLATES["out_punch"].format(name_clean, time_str)
                
                print(f"Applying strict real-time delivery rules. Buffering payload state...")
                await asyncio.sleep(5)
                
                print(f"Dispatching payload data to target: {phone_clean}")
                target_url = f"https://whatsapp.com/send?phone={phone_clean}&text={urllib.parse.quote(status_msg)}"
                
                await page.goto(target_url, {'waitUntil': 'networkidle2'})
                await asyncio.sleep(8)  
                
                await page.evaluate("""() => {
                    const sendBtn = document.querySelector('span[data-icon="send"]') || document.querySelector('button[aria-label="Send"]');
                    if(sendBtn) sendBtn.click();
                }""")
                await asyncio.sleep(2)
                print(f"Message complete for ID: {emp_code}")
                
                last_processed_id = t_id
                
            await asyncio.sleep(2)
            
        except psycopg2.DatabaseError as db_err:
            print(MSG_TEMPLATES["db_err"].format(db_err))
            conn.rollback()
            await asyncio.sleep(10)
        except Exception as loop_err:
            print(MSG_TEMPLATES["sys_err"].format(loop_err))
            await asyncio.sleep(10)
            
    cursor.close()
    conn.close()
    await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
