import os
import shutil
import tempfile
import ctypes
import platform
import hashlib
from pathlib import Path


def clean_temp():
    temp_path = tempfile.gettempdir()
    print(f"\n🧹 Cleaning temporary files in: {temp_path}")
    for filename in os.listdir(temp_path):
        file_path = os.path.join(temp_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f"⚠️ Failed to delete {file_path}: {e}")
    print("✅ Temporary files cleaned.\n")


def clean_recycle_bin():
    print("🗑 Emptying Recycle Bin...")
    if platform.system() == "Windows":
        try:
            ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, 0x00000001)
            print("✅ Recycle Bin emptied.\n")
        except Exception as e:
            print(f"⚠️ Failed to empty Recycle Bin: {e}")
    else:
        print("⚠️ Recycle Bin clean-up is only available on Windows.\n")


def clean_browser_cache():
    print("🌐 Cleaning browser cache...")
    user_profile = Path.home()

    paths = {
        "Chrome": user_profile / "AppData/Local/Google/Chrome/User Data/Default/Cache",
        "Edge": user_profile / "AppData/Local/Microsoft/Edge/User Data/Default/Cache",
        "Firefox": user_profile / "AppData/Local/Mozilla/Firefox/Profiles",
    }

    for browser, path in paths.items():
        if path.exists():
            try:
                shutil.rmtree(path)
                print(f"✅ {browser} cache cleaned.")
            except Exception as e:
                print(f"⚠️ Failed to clean {browser} cache: {e}")
        else:
            print(f"ℹ️ {browser} cache path not found.")
    print()


def scan_duplicates(start_path):
    print(f"🔍 Scanning for duplicate files in {start_path}...")
    hashes = {}
    for root, dirs, files in os.walk(start_path):
        for file in files:
            filepath = os.path.join(root, file)
            try:
                with open(filepath, "rb") as f:
                    file_hash = hashlib.md5(f.read()).hexdigest()
                if file_hash in hashes:
                    print(f"🧭 DUPLICATE: {filepath} == {hashes[file_hash]}")
                else:
                    hashes[file_hash] = filepath
            except:
                continue
    print()


def main():
    print("🚀 Starting Deep Clean Utility...\n")
    clean_temp()
    clean_recycle_bin()
    clean_browser_cache()
    scan_duplicates("C:\\Users\\")  # You can change this path

    print("✅ Deep Clean complete!")


if __name__ == "__main__":
    main()
