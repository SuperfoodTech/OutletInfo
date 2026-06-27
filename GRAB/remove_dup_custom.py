import os
import pandas as pd
import glob

INPUT_FOLDER = os.path.join(os.path.dirname(__file__), "hasil_custom")


def remove_duplicates_from_files():
    """Menghapus baris duplikat dari setiap file Excel di folder hasil_custom."""

    pattern = os.path.join(INPUT_FOLDER, "*.xlsx")
    all_files = [f for f in glob.glob(pattern) if "MASTER" not in os.path.basename(f) and "DUPLIKAT" not in os.path.basename(f)]

    if not all_files:
        print(f"[!] Tidak ada file Excel ditemukan di: {INPUT_FOLDER}")
        return

    print(f"[*] Memproses {len(all_files)} file...\n")

    for filepath in sorted(all_files):
        name = os.path.splitext(os.path.basename(filepath))[0]
        try:
            df = pd.read_excel(filepath)
            original_len = len(df)
            df.dropna(how="all", inplace=True)

            # Tentukan kolom kunci
            key_cols = [col for col in ["Merchant ID", "Item ID", "Nama Item"] if col in df.columns]

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
