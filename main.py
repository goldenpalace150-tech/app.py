# main.py - Deploy this on Render / Railway
from fastapi import FastAPI, Request, Response
import psycopg2
from datetime import datetime
import zoneinfo
import os

app = FastAPI()

# Fetch database credentials securely from your hosting service environment variables
DATABASE_URL = os.environ.get("NEON_DATABASE_URL")
SYRIA_TZ = zoneinfo.ZoneInfo("Asia/Damascus")

@app.get("/")
def home():
    return {"status": "BioTime Listener Online"}

# --- 📠 MAIN ADMS HEARTBEAT HANDSHAKE ENDPOINT ---
@app.get("/iclock/getrequest")
@app.get("/iclock/cdata")
async def adms_gateway(request: Request):
    params = dict(request.query_params)
    # Search for Serial Number safely in uppercase or lowercase strings
    device_sn = params.get("SN") or params.get("sn")
    
    if device_sn:
        now_time = datetime.now(SYRIA_TZ).replace(tzinfo=None)
        try:
            conn = psycopg2.connect(DATABASE_URL, sslmode='require')
            cursor = conn.cursor()
            
            # Instantly update device state inside Neon DB
            cursor.execute("""
                UPDATE iclock_terminal 
                SET last_activity = %s 
                WHERE sn = %s;
            """, (now_time, device_sn))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            # Return raw validation text expected by ZKTeco hardware firmware
            return Response(content="OK", media_type="text/plain")
        except Exception as e:
            return Response(content=f"DB_ERR: {str(e)}", media_type="text/plain")
            
    return Response(content="NO_SN", media_type="text/plain")

# --- 📥 DATA POST RECEIVER FOR PUNCH TRANSFERS ---
@app.post("/iclock/cdata")
async def receive_punches(request: Request):
    # This handles the raw attendance logs sent by the devices
    return Response(content="OK", media_type="text/plain")
