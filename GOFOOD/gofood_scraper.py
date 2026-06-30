import os
import json
import time
import getpass
import urllib.request
import csv
import requests
from pathlib import Path
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv

load_dotenv()

# --- CONFIG ---
APP_SCRIPT_URL = os.getenv("APP_SCRIPT_URL", "")
# --------------

SESSION_DIR = Path(__file__).parent / "session"
SESSION_DIR.mkdir(parents=True, exist_ok=True)

def get_credentials_from_sheet():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRYSUnKOqk29LCktTxdb0wPLbWMbRaWRP3eC_UA4AwYod1FW6zDMhtLMC5ghIvot2B8upCDfBsn-TCP/pub?gid=0&single=true&output=csv"
    try:
        print("[*] Mengambil data portal dari Google Sheet...")
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            lines = [l.decode('utf-8') for l in response.readlines()]
            reader = csv.reader(lines)
            data = list(reader)
            
            portals = []
            for row in data[2:]:
                if len(row) > 5 and row[1].strip():
                    portal = row[1].strip()
                    email = row[3].strip()
                    password = row[5].strip()
                    if email and password:
                        portals.append({
                            'portal': portal,
                            'email': email,
                            'password': password
                        })
            return portals
    except Exception as e:
        print(f"⚠️ Gagal mengambil data dari Google Sheet: {e}")
        return []

def load_session(identifier):
    sanitized = "".join(c for c in identifier if c.isalnum() or c in "._-@")
    session_file = SESSION_DIR / f"session_{sanitized}.json"
    if session_file.exists():
        with open(session_file, 'r') as f:
            return json.load(f)
    return None

def save_session(identifier, session_data):
    sanitized = "".join(c for c in identifier if c.isalnum() or c in "._-@")
    session_file = SESSION_DIR / f"session_{sanitized}.json"
    with open(session_file, 'w') as f:
        json.dump(session_data, f, indent=4)

def parse_selection(choice, max_val):
    selected = set()
    choice = choice.strip()
    if choice.lower() in ('0', 'all'):
        return list(range(max_val))
        
    parts = choice.split(',')
    for part in parts:
        part = part.strip()
        if '-' in part:
            try:
                start, end = map(int, part.split('-'))
                if 1 <= start <= end <= max_val:
                    for i in range(start, end + 1):
                        selected.add(i - 1)
            except ValueError:
                pass
        else:
            try:
                val = int(part)
                if 1 <= val <= max_val:
                    selected.add(val - 1)
            except ValueError:
                pass
    return sorted(list(selected))

def combine_master(output_dir):
    import pandas as pd
    import glob
    import shutil
    import datetime

    print("\n[*] Menggabungkan semua file ke 0master.xlsx...")
    files = glob.glob(os.path.join(output_dir, "GOFOOD_outlets_*.xlsx"))
    
    if not files:
        print("   ⚠️ Tidak ada file hasil scraping untuk digabung.")
        return

    all_dfs = []
    for f in files:
        if os.path.basename(f) == "0master.xlsx":
            continue
        try:
            df = pd.read_excel(f)
            all_dfs.append(df)
        except Exception as e:
            print(f"   ⚠️ Gagal membaca {f}: {e}")
            
    if all_dfs:
        master_df = pd.concat(all_dfs, ignore_index=True)
        master_path = os.path.join(output_dir, "0master.xlsx")
        
        # File versioning
        if os.path.exists(master_path):
            version_dir = os.path.join(output_dir, "versions")
            os.makedirs(version_dir, exist_ok=True)
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = os.path.join(version_dir, f"0master_v{timestamp}.xlsx")
            shutil.copy2(master_path, backup_path)
            print(f"   💾 Master lama di-backup ke: {backup_path}")
            
        master_df.to_excel(master_path, index=False)
        print(f"   ✅ Penggabungan berhasil! Total {len(master_df)} baris disimpan di 0master.xlsx")
        
        # Upload ke Drive jika URL tersedia
        if APP_SCRIPT_URL:
            upload_to_drive(master_path)
        else:
            print("   ℹ️ Upload ke Google Drive dilewati (APP_SCRIPT_URL masih kosong).")

def upload_to_drive(file_path):
    import base64
    print(f"\n[*] Mengunggah {os.path.basename(file_path)} ke Google Drive via Apps Script...")
    try:
        with open(file_path, "rb") as f:
            file_data = f.read()
            
        b64_data = base64.b64encode(file_data).decode('utf-8')
        
        payload = {
            "fileName": os.path.basename(file_path),
            "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "fileData": b64_data
        }
        
        response = requests.post(APP_SCRIPT_URL, json=payload, allow_redirects=True)
        
        if response.status_code in (200, 201):
            try:
                res_json = response.json()
                if res_json.get("status") == "success":
                    print(f"   ✅ Berhasil diunggah! URL: {res_json.get('url')}")
                else:
                    print(f"   ⚠️ Gagal: {res_json.get('message')}")
            except:
                print("   ✅ Berhasil diunggah!")
        else:
            print(f"   ⚠️ Gagal mengunggah. Status code: {response.status_code}")
            print(f"   Detail: {response.text}")
    except Exception as e:
        print(f"   ⚠️ Terjadi kesalahan saat mengunggah: {e}")

def main():
    print("="*60)
    print("  🚀 GOFOOD OUTLET SCRAPER")
    print("="*60)
    
    portals = get_credentials_from_sheet()
    target_portals = []
    
    if portals:
        print("\nDaftar Portal dari Google Sheet:")
        print("  [0] Pilih Semua (All)")
        for idx, p in enumerate(portals):
            print(f"  [{idx+1}] {p['portal']} ({p['email']})")
        print(f"  [n] Input Email Baru secara manual")
        
        choice = input(f"\nPilih portal (contoh: 1, 2-3, all, n): ").strip().lower()
        if choice == 'n':
            email = input("\nMasukkan Email Akun GoFood: ").strip()
            if email:
                target_portals.append({
                    'email': email,
                    'password': "",
                    'portal': "Manual Input"
                })
        else:
            selected_indices = parse_selection(choice, len(portals))
            for idx in selected_indices:
                target_portals.append(portals[idx])
                
    if not target_portals:
        email = input("\nMasukkan Email Akun GoFood: ").strip()
        if email:
            target_portals.append({
                'email': email,
                'password': "",
                'portal': "Manual Input"
            })
            
    if not target_portals:
        print("Tidak ada portal yang dipilih.")
        return

    with sync_playwright() as p:
        print("\n[*] Membuka browser...")
        browser = p.chromium.launch(
            headless=False, 
            args=[
                '--disable-blink-features=AutomationControlled',
                '--disable-infobars',
                '--no-sandbox'
            ]
        )
        
        for target in target_portals:
            email = target['email']
            sheet_password = target.get('password', '')
            portal_name_str = target.get('portal', 'Manual Input')
            
            print(f"\n{'='*50}")
            print(f"  Memproses: {portal_name_str} ({email})")
            print(f"{'='*50}")

            session_data = load_session(email)
            
            context = browser.new_context(
                viewport={'width': 1366, 'height': 768},
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            )

            if session_data and session_data.get('cookies'):
                print("   🔑 Sesi ditemukan. Memuat cookies...")
                context.add_cookies(session_data['cookies'])

            page = context.new_page()
            
            print("[*] Mengakses https://portal.gofoodmerchant.co.id/dashboard ...")
            page.goto("https://portal.gofoodmerchant.co.id/dashboard", wait_until="domcontentloaded")
            
            try:
                page.wait_for_url(lambda url: "/auth" in url or "login" in url or "/gofood" in url, timeout=10000)
            except:
                pass
                
            time.sleep(3)

            # Proses Login
            if "/auth" in page.url or "login" in page.url:
                print("\n⚠️ Memulai proses login otomatis dengan Email & Password...")
                page.goto("https://portal.gofoodmerchant.co.id/auth/login/email", wait_until="domcontentloaded")
                time.sleep(2)
                
                try:
                    email_input = page.wait_for_selector(
                        'input[type="email"], input[name="email"], input[name="username"], input[placeholder*="email" i], input[placeholder*="Email" i]',
                        state='visible',
                        timeout=10000
                    )
                    
                    if email_input:
                        try:
                            email_input.click(force=True, timeout=3000)
                        except:
                            email_input.evaluate("el => el.focus()")
                        time.sleep(0.3)
                        email_input.fill(email, force=True)
                        time.sleep(0.5)
                        
                        pass_input = page.locator('input[type="password"], input[name="password"]')
                        if pass_input.count() == 0:
                            submit_btn = page.locator('button:has-text("Lanjut"), button:has-text("Submit"), button[type="submit"]')
                            if submit_btn.count() > 0:
                                submit_btn.first.click()
                            else:
                                email_input.press("Enter")
                            
                            time.sleep(2)
                            pass_input = page.locator('input[type="password"], input[name="password"]')
                            
                        if pass_input.count() > 0:
                            if sheet_password:
                                password = sheet_password
                                print(f"   🔑 Menggunakan password dari Google Sheet untuk {email}.")
                            else:
                                password = getpass.getpass(f"\n🔑 Masukkan Password untuk {email}: ").strip()
                                
                            if not password:
                                print("⚠️ Password kosong. Menghentikan login.")
                                context.close()
                                continue
                                
                            print("   🤖 Memasukkan password...")
                            try:
                                pass_input.first.click(force=True, timeout=3000)
                            except:
                                pass_input.first.evaluate("el => el.focus()")
                            time.sleep(0.3)
                            pass_input.first.fill(password, force=True)
                            time.sleep(0.5)
                            
                            pass_submit = page.locator('button:has-text("Masuk"), button:has-text("Lanjut"), button[type="submit"]')
                            if pass_submit.count() > 0:
                                pass_submit.first.click(force=True)
                            else:
                                pass_input.first.press("Enter")
                            print("   ✅ Password berhasil diisi. Mengirim form...")
                        else:
                            print("   ⚠️ Kolom password tidak ditemukan.")
                except Exception as e:
                    print(f"   ⚠️ Terjadi kesalahan saat automasi login: {e}")
                    
                print("\n[*] Menunggu login selesai (Otomatis/Manual)...")
                try:
                    page.wait_for_url(lambda url: "/auth/login" not in url, timeout=60000)
                    print("\n✅ Login terdeteksi berhasil!")
                except Exception as e:
                    print("\n❌ Waktu login habis atau browser ditutup.")
                    context.close()
                    continue

            # Simpan sesi
            save_session(email, {
                'email': email,
                'cookies': context.cookies(),
                'timestamp': time.time()
            })
            print("   💾 Sesi login berhasil disimpan.")

            # --- Objektif: Masuk ke portal /outlets ---
            valid_token = None
            def capture_headers(request):
                nonlocal valid_token
                if "api.gobiz.co.id" in request.url or "api.gojekapi.com" in request.url:
                    h = request.headers
                    if 'authorization' in h:
                        valid_token = h['authorization']
            
            page.on("request", capture_headers)

            print("\n[*] Mengakses halaman Outlets...")
            page.goto("https://portal.gofoodmerchant.co.id/outlets", wait_until="domcontentloaded")
            
            print("   [*] Melakukan double refresh halaman untuk memastikan data termuat...")
            time.sleep(2)
            page.reload(wait_until="domcontentloaded")
            time.sleep(2)
            page.reload(wait_until="domcontentloaded")
            time.sleep(3) # Tunggu network selesai agar token tertangkap
            
            print("\n✅ Berhasil masuk ke halaman Outlets (https://portal.gofoodmerchant.co.id/outlets)")
            
            print("\n[*] Mengambil data outlet (Scraping API)...")
            try:
                # 1. Ambil token dari interceptor, fallback ke localStorage
                token = valid_token
                if not token:
                    token = page.evaluate("""() => {
                        const keys = ['token', 'access_token', 'accessToken', 'auth_token', 'authorization', 'gobiz-token', 'go-id-token'];
                        for (const k of keys) {
                            let val = localStorage.getItem(k) || sessionStorage.getItem(k);
                            if (val) {
                                if (val.startsWith('{')) {
                                    try {
                                        const parsed = JSON.parse(val);
                                        val = parsed.token || parsed.access_token || parsed.accessToken || val;
                                    } catch(e){}
                                }
                                if (val && val.length > 20) return val;
                            }
                        }
                        const tokenRegex = /[A-Za-z0-9-_=]+\\.[A-Za-z0-9-_=]+\\.?[A-Za-z0-9-_.+/=]*/;
                        for (let i = 0; i < localStorage.length; i++) {
                            const val = localStorage.getItem(localStorage.key(i));
                            if (val && val.length > 20) {
                                if (val.includes('eyJ')) return val;
                                const match = val.match(tokenRegex);
                                if (match) return match[0];
                            }
                        }
                        for (let i = 0; i < sessionStorage.length; i++) {
                            const val = sessionStorage.getItem(sessionStorage.key(i));
                            if (val && val.length > 20) {
                                if (val.includes('eyJ')) return val;
                                const match = val.match(tokenRegex);
                                if (match) return match[0];
                            }
                        }
                        return null;
                    }""")
                
                if token:
                    if not token.startswith("Bearer "):
                        token = "Bearer " + token
                        
                    payload_str = json.dumps({
                        "from": 0,
                        "size": 1000,
                        "_source": [
                            "id", "director_name", "merchant_name", "email", "feature_types", 
                            "phone", "outlet_address", "outlet_name", "outlet_city", 
                            "payment_settings.GOPAY", "tags", "bank_account", "applications", 
                            "pops", "aspi", "business_type", "metadata", "id_type", "merchant_type"
                        ]
                    })
                    
                    print("   [*] Mengeksekusi request ke GoBiz API...")
                    api_response = page.evaluate("""async ({token, payload}) => {
                        try {
                            const res = await fetch('https://api.gobiz.co.id/v1/merchants/search', {
                                method: 'POST',
                                headers: {
                                    'Accept': 'application/json, text/plain, */*',
                                    'Authentication-Type': 'go-id',
                                    'Authorization': token,
                                    'Content-Type': 'application/json'
                                },
                                body: payload
                            });
                            return await res.json();
                        } catch (e) {
                            return { error: e.message };
                        }
                    }""", {"token": token, "payload": payload_str})
                    
                    outlets_data = []
                    hits = []
                    if api_response and 'hits' in api_response:
                        hits = api_response['hits']
                    elif api_response and 'data' in api_response:
                        hits = api_response['data']
                        
                    if hits:
                        print(f"   ✅ Berhasil menarik {len(hits)} data outlet mentah!")
                        for item in hits:
                            src = item.get('_source', item)
                            
                            portal_name = portal_name_str
                            nama = src.get('outlet_name') or src.get('merchant_name', 'Unknown')
                            store_id = src.get('id', '')
                            
                            # Extract Status
                            status = "Unknown"
                            apps = src.get('applications', {})
                            if 'goresto' in apps:
                                status = apps['goresto'].get('status', status)
                                
                            alamat = src.get('outlet_address', '')
                            
                            # Extract Group ID (if available)
                            group_id = apps.get('goresto', {}).get('goresto_id', store_id)
                            
                            # Extract Bank Account
                            bank_acc = ""
                            if 'bank_account' in src and isinstance(src['bank_account'], dict):
                                bank_no = src['bank_account'].get('account_no', '')
                                bank_name = src['bank_account'].get('bank_name', '')
                                acc_name = src['bank_account'].get('account_name', '')
                                
                                parts = []
                                if bank_name: parts.append(bank_name)
                                if bank_no: parts.append(bank_no)
                                if acc_name: parts.append(acc_name)
                                
                                bank_acc = " - ".join(parts)
                                    
                            outlets_data.append({
                                'Portal': portal_name,
                                'Nama': nama,
                                'Group ID': group_id,
                                'Store ID': store_id,
                                'Status': status,
                                'Alamat': alamat,
                                'Bank Account': bank_acc
                            })
                            
                        # Save to Excel
                        if outlets_data:
                            import pandas as pd
                            df = pd.DataFrame(outlets_data)
                            output_dir = Path(__file__).parent / "data"
                            output_dir.mkdir(exist_ok=True)
                            
                            sanitized_email = "".join(c for c in email if c.isalnum() or c in "._-@")
                            out_file = output_dir / f"GOFOOD_outlets_{sanitized_email}.xlsx"
                            
                            df.to_excel(out_file, index=False)
                            print(f"   💾 Data berhasil disimpan ke: {out_file}")
                            print(f"   📊 Total baris: {len(df)}")
                    else:
                        print(f"   ⚠️ Gagal membaca struktur data API: {api_response.get('error', api_response)}")
                else:
                    print("   ⚠️ Token otentikasi tidak ditemukan di browser.")
            except Exception as e:
                print(f"   ⚠️ Terjadi kesalahan saat scraping: {e}")
                
            print("\n⏳ Proses untuk akun ini selesai. Menutup tab...")
            time.sleep(1)

            try:
                if not page.is_closed():
                    save_session(email, {'email': email, 'cookies': context.cookies(), 'timestamp': time.time()})
                    print("   💾 Sesi terakhir disimpan.")
                context.close()
            except:
                pass
                
        print("\n✅ Semua portal yang dipilih telah selesai diproses.")
        browser.close()
        
    # Lakukan penggabungan master jika seluruh portal dipilih
    is_all_selected = len(target_portals) == len(portals) and len(portals) > 0
    if is_all_selected:
        output_dir = Path(__file__).parent / "data"
        combine_master(str(output_dir))

if __name__ == "__main__":
    main()
