import os
import time

import pythoncom
import win32com.client as win32
import sys


def convert_all_docs_to_docx(source_folder, dest_folder, retries=6):
    os.makedirs(dest_folder, exist_ok=True)

    for filename in os.listdir(source_folder):
        if filename.lower().endswith('.doc') and not filename.lower().endswith('.docx'):
            source_path = os.path.join(source_folder, filename)
            base_name = os.path.splitext(filename)[0]
            dest_path = os.path.join(dest_folder, base_name + ".docx")

            for attempt in range(1, retries + 1):
                try:
                    pythoncom.CoInitialize()  # Required for COM in this thread

                    word = win32.gencache.EnsureDispatch('Word.Application')
                    word.Visible = False

                    doc = word.Documents.Open(source_path)
                    time.sleep(1)  # Let Word stabilize

                    doc.SaveAs(dest_path, FileFormat=16)  # wdFormatDocumentDefault
                    doc.Close(False)
                    word.Quit()
                    print(f"✅ Converted: {filename}")
                    break  # Success, exit retry loop
                except Exception as e:
                    print(f"❌ Attempt {attempt} failed for {filename}: {e}")
                    time.sleep(5)  # Wait before retrying
                    try:
                        word.Quit()
                    except:
                        pass
                finally:
                    pythoncom.CoUninitialize()


if __name__ == "__main__":
    source_path = sys.argv[1]
    dest_path = sys.argv[2]
    sys.exit(convert_all_docs_to_docx(source_path, dest_path))
