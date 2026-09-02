import logging
import shutil
import datetime
import os
import sys

# Pindah ke direktori GRAB agar import berhasil
os.chdir(os.path.join(os.path.dirname(__file__), "GRAB"))
sys.path.append(".")

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger("PostProcessing")

try:
    from combine_custom import combine_excel_files
    from find_duplicates import find_duplicates
    from remove_dup_custom import remove_duplicates_from_files
    from format_excel import apply_formatting_and_sheets
    
    logger.info("--- 1. Menggabungkan file ---")
    combine_excel_files()
    
    timestamp = datetime.datetime.now().strftime("%Y%m%d %H%M")
    raw_filename = f"GRAB {timestamp}.xlsx"
    try:
        apply_formatting_and_sheets("hasil_custom/0MASTER.xlsx", f"hasil_custom/{raw_filename}", create_tabs=False)
        logger.info(f"Data mentah disimpan dengan format: hasil_custom/{raw_filename}")
    except Exception as ex:
        logger.error(f"Gagal menyimpan {raw_filename}: {ex}")
    
    logger.info("--- 2. Mencari duplikat ---")
    find_duplicates()
    
    logger.info("--- 3. Menghapus duplikat dari file individual ---")
    remove_duplicates_from_files()
    
    logger.info("--- 4. Menggabungkan kembali file yang sudah bersih ---")
    combine_excel_files()
    
    processed_filename = f"GRAB PROCESSED {timestamp}.xlsx"
    try:
        apply_formatting_and_sheets("hasil_custom/0MASTER.xlsx", f"hasil_custom/{processed_filename}", create_tabs=True)
        logger.info(f"Data bersih disimpan dengan format dan tabs: hasil_custom/{processed_filename}")
    except Exception as ex:
        logger.error(f"Gagal menyimpan {processed_filename}: {ex}")
    
    logger.info("Semua proses post-processing selesai!")
except Exception as e:
    logger.error(f"Error during post-processing: {e}")

