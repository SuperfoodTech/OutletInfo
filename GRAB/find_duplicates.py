import os
import pandas as pd
import glob

INPUT_FOLDER = os.path.join(os.path.dirname(__file__), "hasil_custom")
MASTER_FILE = os.path.join(INPUT_FOLDER, "MASTER_ALL.xlsx")


def find_duplicates():
    """Mencari baris duplikat di file master dan menyimpan hasilnya."""

    if not os.path.exists(MASTER_FILE):
        print(f"[!] File master tidak ditemukan: {MASTER_FILE}")
        print("    Jalankan combine_custom.py terlebih dahulu.")
        return

    df = pd.read_excel(MASTER_FILE)
    print(f"[*] Memuat {len(df)} baris dari {os.path.basename(MASTER_FILE)}")

    # Tentukan kolom kunci untuk deteksi duplikat
    key_cols = []
    for col in ["Merchant ID", "Item ID", "Nama Item"]:
        if col in df.columns:
            key_cols.append(col)

    if not key_cols:
        print("[!] Tidak ada kolom kunci (Merchant ID / Item ID / Nama Item) ditemukan.")
        return

    print(f"[*] Menggunakan kolom kunci: {key_cols}")

    # Tandai duplikat
    duplicates_mask = df.duplicated(subset=key_cols, keep=False)
    duplicates_df = df[duplicates_mask].copy()

    if duplicates_df.empty:
        print("[✓] Tidak ada duplikat ditemukan.")
        return

    print(f"\n[!] Ditemukan {len(duplicates_df)} baris duplikat ({duplicates_df[key_cols].drop_duplicates().shape[0]} grup unik):")

    # Tampilkan ringkasan
    for keys, group in duplicates_df.groupby(key_cols):
        label = " | ".join(str(k) for k in (keys if isinstance(keys, tuple) else [keys]))
        print(f"    - {label}: {len(group)} kemunculan")

    # Simpan file duplikat
    out_file = os.path.join(INPUT_FOLDER, "DUPLIKAT.xlsx")
    duplicates_df.to_excel(out_file, index=False)
    print(f"\n[✓] Data duplikat disimpan ke: {out_file}")


if __name__ == "__main__":
    find_duplicates()
