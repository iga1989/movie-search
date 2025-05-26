import shutil
import time
import pythoncom
import os
import win32com.client as win32


def copy_files_from_deep_folders(current_path, destination_dir):
    if not os.path.isdir(current_path):
        return

    entries = os.listdir(current_path)

    # Get all files in the current directory
    files = [f for f in entries if os.path.isfile(os.path.join(current_path, f))]
    if files:
        for file in files:
            src_path = os.path.join(current_path, file)
            dst_path = os.path.join(destination_dir, file)

            # Handle duplicate file names
            base, ext = os.path.splitext(file)
            counter = 1
            while os.path.exists(dst_path):
                dst_path = os.path.join(destination_dir, f"{base}_{counter}{ext}")
                counter += 1

            shutil.copy2(src_path, dst_path)
            print(f"Copied: {src_path} -> {dst_path}")
    else:
        # No files found, go deeper
        subdirs = [d for d in entries if os.path.isdir(os.path.join(current_path, d))]
        for subdir in subdirs:
            copy_files_from_deep_folders(os.path.join(current_path, subdir), destination_dir)



def convert_all_docs_to_docx(source_folder, dest_folder, retries=5):
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
                    time.sleep(2)  # Let Word stabilize

                    doc.SaveAs(dest_path, FileFormat=16)  # wdFormatDocumentDefault
                    doc.Close(False)
                    word.Quit()
                    print(f"✅ Converted: {filename}")
                    break  # Success, exit retry loop
                except Exception as e:
                    print(f"❌ Attempt {attempt} failed for {filename}: {e}")
                    time.sleep(1)  # Wait before retrying
                    try:
                        word.Quit()
                    except:
                        pass
                finally:
                    pythoncom.CoUninitialize()

# Example usage:
source = r'C:\Users\isset\Downloads\destination_1'
dest = r'C:\Users\isset\Downloads\destination_3'
# convert_all_docs_to_docx(source, dest)

# Set the paths
source_root = r'C:\Users\isset\Downloads\destination_3'
destination_folder = r'C:\Users\isset\Downloads\destination'
destination_folder_1 = r'C:\Users\isset\Downloads\destination_1'
destination_folder_2 = r'C:\Users\isset\Downloads\destination_2'

# Ensure the destination folder exists
os.makedirs(destination_folder, exist_ok=True)

# Start copying
# copy_files_from_deep_folders(source_root, destination_folder)
# convert_doc_to_docx(destination_folder_1, destination_folder_2)
print("All files copied.")

if __name__ == "__main__":
    src = r'C:\Users\isset\Downloads\hansards_2025'
    dst = r'C:\Users\isset\Downloads\destination_4'

    start_time = time.time()
    copy_files_from_deep_folders(src, dst)
    end_time = time.time()

    print(f"\n✅ Finished in {end_time - start_time:.2f} seconds.")