import os
import pandas as pd
import glob

# Folder tempat file Excel hasil scraping per outlet
INPUT_FOLDER = os.path.join(os.path.dirname(__file__), "hasil_custom")
OUTPUT_FILE = os.path.join(INPUT_FOLDER, "MASTER_ALL.xlsx")


def combine_excel_files():
    """Menggabungkan semua file Excel di folder hasil_custom menjadi satu file master."""
    
    # Cari semua file Excel kecuali file MASTER itu sendiri
    pattern = os.path.join(INPUT_FOLDER, "*.xlsx")
    all_files = [f for f in glob.glob(pattern) if "MASTER" not in os.path.basename(f)]

    if not all_files:
        print(f"[!] Tidak ada file Excel ditemukan di: {INPUT_FOLDER}")
        return

    print(f"[*] Ditemukan {len(all_files)} file:")
    for f in sorted(all_files):
        print(f"    - {os.path.basename(f)}")

    dfs = []
    for filepath in sorted(all_files):
        name = os.path.splitext(os.path.basename(filepath))[0]
        try:
            df = pd.read_excel(filepath)
            # Hapus baris yang sepenuhnya kosong
            df.dropna(how="all", inplace=True)
            # Tambah kolom sumber file jika belum ada
            if "Sumber" not in df.columns:
                df.insert(0, "Sumber", name)
            print(f"    [+] {name}: {len(df)} baris")
            dfs.append(df)
        except Exception as e:
            print(f"    [!] Gagal membaca {name}: {e}")

    if not dfs:
        print("[!] Tidak ada data yang berhasil dimuat.")
        return

    master_df = pd.concat(dfs, ignore_index=True)

    # Hapus lagi baris kosong setelah penggabungan
    master_df.dropna(how="all", inplace=True)
    master_df.reset_index(drop=True, inplace=True)

    master_df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n[✓] Master file disimpan: {OUTPUT_FILE}")
    print(f"    Total baris: {len(master_df)}")
    print(f"    Total kolom: {len(master_df.columns)}")


if __name__ == "__main__":
    combine_excel_files()
