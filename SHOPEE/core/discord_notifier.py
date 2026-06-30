import os
import json
import time
from pathlib import Path

def send_discord_error(platform: str, merchant: str, error_type: str, message: str, phone: str = ""):
    """
    Menyimpan informasi error ke folder data/discord_notifications 
    agar Discord Bot dapat memantaunya dan mengirim notifikasi cantik.
    
    error_type: 'OTP_TIMEOUT', 'WRONG_CREDENTIALS', 'BLOCKED_ACCOUNT', 'SYSTEM_ERROR' dll.
    message: Pesan ramah manusia (bukan teknis)
    """
    script_dir = Path(__file__).resolve().parent.parent
    notif_dir = script_dir / "data" / "discord_notifications"
    notif_dir.mkdir(parents=True, exist_ok=True)
    
    timestamp = int(time.time() * 1000)
    filename = f"error_{platform}_{timestamp}.json"
    filepath = notif_dir / filename
    
    payload = {
        "status": "ERROR_NOTIF",
        "platform": platform,
        "merchant": merchant,
        "error_type": error_type,
        "message": message,
        "phone": phone,
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "channel_id": os.environ.get("OFD_CHANNEL_ID", "")
    }
    
    try:
        with open(filepath, "w") as f:
            json.dump(payload, f, indent=2)
        print(f"✅ Notifikasi error telah disimpan untuk Discord: {filename}")
    except Exception as e:
        print(f"⚠️ Gagal menyimpan notifikasi Discord: {e}")
