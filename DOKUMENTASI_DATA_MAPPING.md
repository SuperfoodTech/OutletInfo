# 📋 Dokumentasi Pemetaan Data & Endpoint API - Outlet Info

Dokumen ini berisi spesifikasi teknis pemetaan **10 Kolom Objective Standard Superfood Tech**, sumber endpoint API platform (*GoFood*, *GrabFood*, dan *ShopeeFood*), serta logika ekstraksi dan deduplikasi data.

---

## 1. Spesifikasi Template Excel (`YYYY-MM-DD HH_MM Nama Pemilik.xlsx`)

Template standar yang digunakan adalah sheet **`Listing`** pada file [`YYYY-MM-DD HH_MM Nama Pemilik.xlsx`](file:///mnt/DATA/Proyek/Outlet%20Info/YYYY-MM-DD%20HH_MM%20Nama%20Pemilik.xlsx) yang memiliki 37 kolom header.

### 📌 Aturan Pengisian Kolom:
* **Hanya mengisi Kolom A, B, dan G sampai Q (13 kolom terisi):**
  * **Kolom A (`Nama Pemilik`):** Nama pemilik / owner portal (`VB` / nama pemilik dari master sheet).
  * **Kolom B (`Nama Brand`):** Nama brand induk toko/portal.
  * *Kolom C - F (`Model`, `Tipe`, `Outlet`, `Nomor HP`): Dikosongkan.*
  * **Kolom G (`Aplikator`):** `GoFood` / `GrabFood` / `ShopeeFood`.
  * **Kolom H (`Nama Portal`):** Nama portal/brand credential dari Google Sheet.
  * **Kolom I (`Group ID`):** Group / Entity ID dari masing-masing aplikator.
  * **Kolom J (`Nama Listing`):** Nama outlet / listing toko.
  * **Kolom K (`Link`):** URL konsumen restoran (`http://gofood.co.id/surabaya/restaurant/{rest_id}`).
  * **Kolom L (`Store ID`):** ID Unik Toko (misal GoBiz `G...`).
  * **Kolom M (`Status Listing`):** Status operasional toko (`active` / `inactive`).
  * **Kolom N (`Alamat`):** Alamat lengkap outlet.
  * **Kolom O (`Nama Bank`):** Nama bank pencairan.
  * **Kolom P (`Nama Pemilik Rekening`):** Nama pemilik rekening.
  * **Kolom Q (`Nomor Rekening`):** Nomor rekening bank (format teks `@`).
  * *Kolom R s/d AK (Kolom 18–37): Dibiarkan kosong sesuai template.*

---

## 2. Pemetaan Sumber Field API Per Aplikator

### A. GoFood (GoBiz API)

* **Metode Login:** Otentikasi Email via OTP Otomatis (Agency) & Kata Sandi/Password (VB / Owner).
* **Pengelompokan & Deduplikasi Scraping:**
  * Kredensial login dideduplikasi per email unik sehingga akun multi-outlet (misal *Holans* dengan 10 baris di Google Sheet) hanya di-login **1 kali**.
  * Akun-akun login dikelompokkan berdasarkan **`Nama Pemilik` (Owner)**. Seluruh outlet dari multi-kredensial milik satu pemilik (misal Owner *AGSA* punya 2 kredensial) otomatis digabungkan ke dalam 1 file output per-owner.
* **Endpoint Utama:** `POST https://api.gobiz.co.id/v1/merchants/search`
* **Headers:**
  * `Authorization`: `Bearer {access_token}`
  * `Authentication-Type`: `go-id`
  * `Content-Type`: `application/json`

#### Pemetaan Field GoFood:

| Kolom Objective | Sumber Field GoBiz API | Contoh Nilai |
| :--- | :--- | :--- |
| `Aplikator` | Statis: `"GoFood"` | `GoFood` |
| `Nama Outlet` | `_source.outlet_name` *(Fallback: `_source.merchant_name`)* | `Ayam Geprek Suroboyo, Ampel` |
| `Nama Portal` | Master Google Sheet / CSV Credential (`Brand`) | `AGSA - Ayam Geprek Suroboyo Ampel` |
| `Group ID` | `_source.external_ids.entity[0]` / `_source.tags.entity[0]` *(Fallback: `partner_id`)* | `001Id000006MCj0IAG` |
| `Link` | `http://gofood.co.id/surabaya/restaurant/{rest_id}` *(Sumber: `applications.goresto.goresto_id` / Restaurant UUID)* | `http://gofood.co.id/surabaya/restaurant/145d7fd4-5168-43ac-b3cf-d8d5a496d3a2` |
| `Store ID` | `_source.id` *(GoBiz Merchant ID berawalan `G...`)* | `G025124092` |
| `Status` | `_source.applications.goresto.status` | `active` / `inactive` |
| `Alamat` | `_source.outlet_address` | `Jl. Nyamplungan No. 123, Surabaya` |
| `Nama Bank` | `_source.bank_account.bank_name` | `BCA` |
| `Nama Pemilik Rekening` | `_source.bank_account.account_name` | `PT SUPERFOOD BERKAH` |
| `Nomor Rekening` | `_source.bank_account.account_no` | `1900313230` |

---

### B. GrabFood (Grab Merchant API)

* **Metode Login:** Otentikasi Username & Password (Playwright dengan simulasi *human typing* dan penanganan *welcome back*).
* **Endpoint Group ID:** `GET https://merchant.grab.com/troy/user-profile/v1/merchant-selector`
* **Endpoint Search Outlet (1-Pass):** `GET https://api.grab.com/delvplatformapi/merchant/v1/merchant-group/store/search?offset=0&limit=100&search=&includeInactive=true&asc=true&cityIDs[]=ALL`
* **Endpoint Fallback Detail Bank:** `GET https://merchant.grab.com/troy/v1/merchant?merchant_group_id={idmg}&isBalanceNeeded=false&currency=IDR` *(Header: `x-mex-resource: zeus_store:{merchant_id}`)*

#### Pemetaan Field GrabFood:

| Kolom Objective | Sumber Field Grab Merchant API | Contoh Nilai |
| :--- | :--- | :--- |
| `Aplikator` | Statis: `"GrabFood"` | `GrabFood` |
| `Nama Outlet` | `stores[].name` *(Fallback: `stores[].merchantName`)* | `Spesial Steak Hauchek - Krembangan` |
| `Nama Portal` | Master Google Sheet / CSV Credential (`Brand` / `Portal`) | `AGSA - Spesial Steak Hauchek` |
| `Group ID` | RegEx `(IDMG\d+)` dari response `/troy/user-profile/v1/merchant-selector` | `IDMG20200909102926236239` |
| `Store ID` | `stores[].merchantID` / `stores[].id` | `6-C8DAAA4VG3E3LX` / `IDGFSTI000038j3` |
| `Status` | `stores[].status` / `stores[].isActive` | `ACTIVE` / `INACTIVE` |
| `Alamat` | `stores[].address` *(Fallback: `AddressLine1`)* | `Jl. Krembangan Barat No. 45` |
| `Nama Bank` | `stores[].bankAccount.bankName` *(Fallback: `bank_details.bank_name`)* | `Mandiri` |
| `Nama Pemilik Rekening` | `stores[].bankAccount.accountHolderName` *(Fallback: `bank_details.account_name`)* | `Superfood Tech` |
| `Nomor Rekening` | `stores[].bankAccount.accountNumber` *(Fallback: `bank_details.account_number`)* | `88201026338530` |

---

## 3. Penjelasan Khusus Mengenai `Group ID`

* **GrabFood (`IDMG`):** Merupakan *Merchant Group ID* di sistem backend Grab yang mengikat multi-cabang di bawah 1 kontrak/akun pemilik.
* **GoFood (`001Id...`):** Merupakan *Gojek Master Entity / Brand Salesforce ID* di sistem backend Gojek (`external_ids.entity` & `external_ids.brand` / `tags.entity`). ID ini memayungi seluruh cabang toko milik brand/perusahaan yang sama.
* **ShopeeFood (`MID / Entity ID`):** Merupakan *Merchant ID* induk di Shopee TOB yang menaungi seluruh cabang branch resto.

---

## 4. Standar Output & Aturan Deduplikasi

1. **Format File Output:**
   * `0master.xlsx` — Data gabungan seluruh outlet bersih tanpa duplikat.
   * `YYYY-MM-DD HH_MM [Nama Pemilik].xlsx` — File hasil gabungan per-pemilik (menggunakan template 37 kolom).
   * `GOFOOD_outlets_[Nama Portal].xlsx` — Cache file hasil scraping per portal.
2. **Aturan Deduplikasi:**
   * Di tingkat pemilik (*Owner*): Seluruh outlet dari multi-kredensial dideduplikasi berdasarkan `Store ID`.
   * Di tingkat master: Seluruh dataset digabungkan dan dibersihkan dengan:
     ```python
     master_df.drop_duplicates(subset=["Store ID"], keep="first", inplace=True)
     ```
   * Menjamin tidak ada outlet/cabang yang tercatat ganda meskipun beberapa portal berbagi akun login atau Merchant Group yang sama.

