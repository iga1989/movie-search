import os
import subprocess
import time
import sys
import win32com.client as win32
sys.coinit_flags = 0
import pythoncom
from multiprocessing import Pool, cpu_count
import unicodedata
import ctypes


# Prevent sleep
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)


def kill_winword():
    # Suppress console output and errors
    subprocess.call('taskkill /f /im winword.exe', shell=True,
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def convert_doc_to_docx(args):
    source_path, dest_path, filename, retries = args
    for attempt in range(1, retries + 1):
        try:
            # kill_winword()  # 🔧 Kill any leftover Word processes
            pythoncom.CoInitialize()
            word = win32.gencache.EnsureDispatch('Word.Application')
            word.Visible = False

            doc = word.Documents.Open(source_path)
            print(f"Trying to open: {source_path}")
            time.sleep(5)  # Let Word stabilize

            doc.SaveAs(dest_path, FileFormat=16)  # wdFormatDocumentDefault
            doc.Close(False)
            word.Quit()
            print(f"✅ Converted: {filename}")
            return
        except Exception as e:
            print(f"❌ Attempt {attempt} failed for {filename}: {e}")
            time.sleep(5)
            try:
                word.Quit()
            except:
                pass
        finally:
            pythoncom.CoUninitialize()


def convert_all_docs_to_docx_parallel(source_folder, dest_folder, retries=6):
    os.makedirs(dest_folder, exist_ok=True)

    tasks = []
    for filename in os.listdir(source_folder):
        if filename.lower().endswith('.doc') and not filename.lower().endswith('.docx'):
            # filename = unicodedata.normalize("NFKD", filename)
            source_path = os.path.join(source_folder, filename)
            base_name = os.path.splitext(filename)[0]
            dest_path = os.path.join(dest_folder, base_name + ".docx")
            tasks.append((source_path, dest_path, filename, retries))

    print(f"🔧 Starting conversion using {cpu_count()} cores...")
    with Pool(processes=cpu_count()) as pool:
        pool.map(convert_doc_to_docx, tasks)

    print("✅ All conversions complete.")


# Example usage
if __name__ == "__main__":
    source_folder = r"C:\Users\isset\Downloads\destination_1"
    dest_folder = r"C:\Users\isset\Downloads\destination_3"
    convert_all_docs_to_docx_parallel(source_folder, dest_folder)
