import os
import subprocess
import time
import shutil
from multiprocessing import Pool, cpu_count

def convert_with_libreoffice(args):
    source_path, output_dir, retries, wait_time, failed_dir = args
    filename = os.path.basename(source_path)
    base_name, _ = os.path.splitext(filename)
    dest_file_path = os.path.join(output_dir, base_name + ".docx")

    if os.path.exists(dest_file_path):
        print(f"⏩ Skipping {filename}, .docx already exists.")
        return

    for attempt in range(1, retries + 1):
        try:
            print(f"🔄 Attempt {attempt}: Converting {filename}")
            subprocess.run([
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                "--headless",
                "--convert-to", "docx",
                "--outdir", output_dir,
                source_path
            ], check=True)
            print(f"✅ Converted: {filename}")
            return
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed attempt {attempt} for {filename}: {e}")
            if attempt < retries:
                time.sleep(wait_time)
            else:
                print(f"🚫 Giving up on {filename} after {retries} attempts.")
                try:
                    os.makedirs(failed_dir, exist_ok=True)
                    shutil.copy2(source_path, os.path.join(failed_dir, filename))
                    print(f"📁 Copied failed file to: {failed_dir}")
                except Exception as copy_err:
                    print(f"⚠️ Failed to copy file {filename} to failed folder: {copy_err}")

def convert_all_docs(source_folder, dest_folder, failed_folder, retries=3, wait_time=5):
    os.makedirs(dest_folder, exist_ok=True)
    os.makedirs(failed_folder, exist_ok=True)

    tasks = []
    for filename in os.listdir(source_folder):
        if filename.lower().endswith('.doc') and not filename.lower().endswith('.docx'):
            base_name, _ = os.path.splitext(filename)
            dest_file_path = os.path.join(dest_folder, base_name + ".docx")

            if os.path.exists(dest_file_path):
                print(f"⏩ Already exists: {dest_file_path}")
                continue

            source_path = os.path.join(source_folder, filename)
            tasks.append((source_path, dest_folder, retries, wait_time, failed_folder))

    print(f"🔧 Starting LibreOffice conversion using {cpu_count()} cores...")
    with Pool(processes=cpu_count()) as pool:
        pool.map(convert_with_libreoffice, tasks)

    print("✅ All conversions complete.")

if __name__ == "__main__":
    source_folder = r"C:\Users\isset\Downloads\destination"
    dest_folder = r"C:\Users\isset\Downloads\destination_3"
    failed_folder = r"C:\Users\isset\Downloads\failed_conversions_1"
    convert_all_docs(source_folder, dest_folder, failed_folder, retries=6, wait_time=2)
