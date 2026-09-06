#!/usr/bin/env python3
"""
pipeline_lock.py
================
Job Lock manager untuk mencegah eksekusi konkruen pada pipeline scraping
(GoFood, GrabFood, ShopeeFood) dan penulisan file Excel.
Menggunakan filelock OS-level locking (fcntl/flock) yang otomatis terlepas
bahkan jika proses mengalami crash atau dihentikan secara paksa.
"""

import os
import json
import datetime
from pathlib import Path
from filelock import FileLock, Timeout

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_LOCK_FILE = BASE_DIR / ".pipeline.lock"
DEFAULT_INFO_FILE = BASE_DIR / ".pipeline.info.json"


def is_pipeline_locked(lock_file=None):
    """
    Memeriksa apakah pipeline saat ini sedang dikunci oleh proses lain.
    Mencoba acquire dengan timeout=0. Jika gagal, berarti sedang ada proses aktif.
    """
    target_lock = Path(lock_file) if lock_file else DEFAULT_LOCK_FILE
    lock = FileLock(str(target_lock))
    try:
        lock.acquire(timeout=0)
        lock.release()
        return False
    except Timeout:
        return True


def get_lock_info(info_file=None):
    """
    Membaca metadata tugas yang saat ini sedang memegang lock.
    """
    target_info = Path(info_file) if info_file else DEFAULT_INFO_FILE
    if target_info.exists():
        try:
            with open(target_info, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


class JobLock:
    """
    Context manager untuk penguncian global pipeline.
    
    Penggunaan:
        with JobLock(timeout=0, task_info={"owner": "...", "user": "..."}):
            # Jalankan scraping / export
    """

    def __init__(self, lock_file=None, info_file=None, timeout=0, task_info=None):
        self.lock_file = Path(lock_file) if lock_file else DEFAULT_LOCK_FILE
        self.info_file = Path(info_file) if info_file else DEFAULT_INFO_FILE
        self.timeout = timeout
        self.task_info = task_info or {}
        self._file_lock = FileLock(str(self.lock_file))

    def acquire(self):
        self._file_lock.acquire(timeout=self.timeout)
        try:
            info = {
                "pid": os.getpid(),
                "started_at": datetime.datetime.now().isoformat(),
                **self.task_info
            }
            with open(self.info_file, "w", encoding="utf-8") as f:
                json.dump(info, f, indent=2)
        except Exception:
            pass

    def release(self):
        try:
            if self.info_file.exists():
                self.info_file.unlink(missing_ok=True)
        except Exception:
            pass
        self._file_lock.release()

    def __enter__(self):
        self.acquire()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.release()

