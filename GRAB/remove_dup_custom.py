import os
import pandas as pd
import glob
import re

INPUT_FOLDER = os.path.join(os.path.dirname(__file__), "hasil_custom")


def remove_duplicates_from_files():
    """Menghapus baris duplikat dari setiap file Excel di folder hasil_custom."""

    pattern = os.path.join(INPUT_FOLDER, "*.xlsx")
    all_files = [f for f in glob.glob(pattern) if "MASTER" not in os.path.basename(f) and "DUPLIKAT" not in os.path.basename(f)]

    if not all_files:
        print(f"[!] Tidak ada file Excel ditemukan di: {INPUT_FOLDER}")
        return

    latest_files = {}
    for filepath in all_files:
        filename = os.path.basename(filepath)
        match = re.match(r"^([A-Za-z0-9]+)(?:_(\d+))?\.xlsx$", filename)
        if match:
            base_name = match.group(1)
            version = int(match.group(2)) if match.group(2) else 1
            if base_name not in latest_files or version > latest_files[base_name]['version']:
                latest_files[base_name] = {'filepath': filepath, 'version': version, 'name': base_name}

    files_to_process = sorted([v['filepath'] for v in latest_files.values()])

    print(f"[*] Memproses {len(files_to_process)} file terbaru...\n")

    for filepath in files_to_process:
        name = os.path.basename(filepath)
        try:
            df = pd.read_excel(filepath)
            original_len = len(df)
            df.dropna(how="all", inplace=True)

            # Tentukan kolom kunci
            key_cols = [col for col in ["Store ID", "Merchant ID", "Item ID", "Nama Item"] if col in df.columns]

            if key_cols:
                df.drop_duplicates(subset=key_cols, keep="first", inplace=True)
            else:
                df.drop_duplicates(keep="first", inplace=True)

            df.reset_index(drop=True, inplace=True)
            removed = original_len - len(df)

            df.to_excel(filepath, index=False)

            if removed > 0:
                print(f"[✓] {name}: {removed} duplikat dihapus ({len(df)} baris tersisa)")
            else:
                print(f"[-] {name}: tidak ada duplikat ({len(df)} baris)")

        except Exception as e:
            print(f"[!] Gagal memproses {name}: {e}")

    print("\n[✓] Selesai. Jalankan combine_custom.py untuk membuat ulang file MASTER.")


if __name__ == "__main__":
    remove_duplicates_from_files()
