import os
import json
import time
import getpass
import urllib.request
import csv
from pathlib import Path
from playwright.sync_api import sync_playwright

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

def main():
    print("="*60)
    print("  🚀 GOFOOD OUTLET SCRAPER")
    print("="*60)
    
    portals = get_credentials_from_sheet()
    email = ""
    sheet_password = ""
    portal_name_str = "Manual Input"
    
    if portals:
        print("\nDaftar Portal dari Google Sheet:")
        for idx, p in enumerate(portals):
            print(f"  [{idx+1}] {p['portal']} ({p['email']})")
        print(f"  [n] Input Email Baru secara manual")
        
        choice = input(f"\nPilih portal (1-{len(portals)}/n): ").strip().lower()
        if choice == 'n':
            email = input("\nMasukkan Email Akun GoFood: ").strip()
        else:
            try:
                idx = int(choice) - 1
                if 0 <= idx < len(portals):
                    email = portals[idx]['email']
                    sheet_password = portals[idx]['password']
                    portal_name_str = portals[idx]['portal']
            except ValueError:
                pass
                
    if not email:
        email = input("\nMasukkan Email Akun GoFood: ").strip()
        
    if not email:
        print("Email tidak boleh kosong.")
        return

    session_data = load_session(email)

    with sync_playwright() as p:
        print("[*] Membuka browser...")
        browser = p.chromium.launch(
            headless=False, 
            args=[
                '--disable-blink-features=AutomationControlled',
                '--disable-infobars',
                '--no-sandbox'
            ]
        )
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
                            return
                            
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
                return

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
            
        print("\n⏳ Proses selesai. Menutup browser dalam 3 detik...")
        time.sleep(3)

        try:
            if not page.is_closed():
                save_session(email, {'email': email, 'cookies': context.cookies(), 'timestamp': time.time()})
                print("   💾 Sesi terakhir disimpan.")
            browser.close()
        except:
            pass

if __name__ == "__main__":
    main()
