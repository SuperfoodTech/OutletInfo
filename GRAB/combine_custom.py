import os
import pandas as pd
import glob
import re

# Folder tempat file Excel hasil scraping per outlet
INPUT_FOLDER = os.path.join(os.path.dirname(__file__), "hasil_custom")
OUTPUT_FILE = os.path.join(INPUT_FOLDER, "0MASTER.xlsx")


def combine_excel_files():
    """Menggabungkan semua file Excel di folder hasil_custom menjadi satu file master."""
    
    # Cari semua file Excel kecuali file MASTER itu sendiri
    pattern = os.path.join(INPUT_FOLDER, "*.xlsx")
    all_files = [f for f in glob.glob(pattern) if "master" not in os.path.basename(f).lower() and "duplikat" not in os.path.basename(f).lower() and not os.path.basename(f).upper().startswith("GRAB ")]

    if not all_files:
        print(f"[!] Tidak ada file Excel ditemukan di: {INPUT_FOLDER}")
        return

    # Group by base name and find the latest
    latest_files = {}
    for filepath in all_files:
        filename = os.path.basename(filepath)
        # Parse base name and version
        match = re.match(r"^([A-Za-z0-9]+)(?:_(\d+))?\.xlsx$", filename)
        if match:
            base_name = match.group(1)
            version = int(match.group(2)) if match.group(2) else 1
            if base_name not in latest_files or version > latest_files[base_name]['version']:
                latest_files[base_name] = {'filepath': filepath, 'version': version, 'name': base_name}

    files_to_combine = sorted([v['filepath'] for v in latest_files.values()])

    print(f"[*] Ditemukan {len(files_to_combine)} file terbaru untuk digabungkan:")
    for f in files_to_combine:
        print(f"    - {os.path.basename(f)}")

    dfs = []
    for filepath in files_to_combine:
        filename = os.path.basename(filepath)
        base_name = None
        for k, v in latest_files.items():
            if v['filepath'] == filepath:
                base_name = k
                break
        
        try:
            df = pd.read_excel(filepath)
            # Hapus baris yang sepenuhnya kosong
            df.dropna(how="all", inplace=True)
            # Tambah kolom sumber file jika belum ada (gunakan base_name agar tetap rapi, misal F1 bukan F1_2)
            if "Sumber" not in df.columns:
                df.insert(0, "Sumber", base_name)
            print(f"    [+] {filename}: {len(df)} baris")
            dfs.append(df)
        except Exception as e:
            print(f"    [!] Gagal membaca {filename}: {e}")

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
