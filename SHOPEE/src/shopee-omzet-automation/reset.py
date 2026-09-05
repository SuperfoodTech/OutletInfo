import os
import shutil
from pathlib import Path
import time

def reset_and_login():
    print("\n" + "=" * 60)
    print("  SHOPEE SESSION RESET & RE-LOGIN TOOL")
    print("=" * 60)

    # 1. Define paths
    base_dir = Path(__file__).parent
    session_file = base_dir / "data" / "session.json"
    profile_dir = base_dir / "data" / "chrome_profile"

    # 2. Delete Session JSON
    if session_file.exists():
        try:
            os.remove(session_file)
            print("✅ Successfully removed data/session.json")
        except Exception as e:
            print(f"⚠️  Warning: Could not remove session file: {e}")
    else:
        print("ℹ️  data/session.json already clean.")

    # 3. Delete Chrome Profile (Cookies/Cache)
    if profile_dir.exists():
        print("⏳ Removing Chrome profile folder (cleaning cookies/cache)...")
        # Retry logic for directory removal (sometimes files are locked for a few ms)
        for i in range(3):
            try:
                shutil.rmtree(profile_dir)
                print("✅ Successfully cleaned data/chrome_profile")
                break
            except Exception as e:
                if i == 2:
                    print(f"❌ Failed to remove profile folder: {e}")
                    print("👉 Please make sure ALL zChrome windows are closed and try again.")
                    return
                time.sleep(1)

    # 4. Restart the pipeline
    print("\n🚀 Restarting Shopee Omzet Pipeline...")
    print("-" * 60)
    # Using python directly to ensure it runs in the same environment
    os.system("uv run run_omzet.py")

if __name__ == "__main__":
    reset_and_login()
