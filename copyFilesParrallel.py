import os
import shutil
import time
from multiprocessing import Pool, cpu_count


def copy_files_in_folder(args):
    current_path, destination_dir = args

    if not os.path.isdir(current_path):
        return

    try:
        entries = os.listdir(current_path)
    except PermissionError:
        return

    files = [f for f in entries if os.path.isfile(os.path.join(current_path, f))]
    for file in files:
        src_path = os.path.join(current_path, file)
        dst_path = os.path.join(destination_dir, file)

        # ❌ Skip if file with same name already exists
        if os.path.exists(dst_path):
            print(f"⚠️ Skipped (duplicate): {dst_path}")
            continue

        try:
            shutil.copy2(src_path, dst_path)
            print(f"📁 Copied: {src_path} ➜ {dst_path}")
        except Exception as e:
            print(f"❌ Failed to copy {src_path}: {e}")


def get_all_folders_with_files(root):
    folder_list = []
    for dirpath, dirnames, filenames in os.walk(root):
        if filenames:
            folder_list.append(dirpath)
    return folder_list


def copy_files_from_deep_folders_parallel(source_root, destination_folder):
    os.makedirs(destination_folder, exist_ok=True)
    folder_list = get_all_folders_with_files(source_root)

    print(f"🧠 Using {cpu_count()} cores to copy from {len(folder_list)} folders...")
    with Pool(processes=cpu_count()) as pool:
        pool.map(copy_files_in_folder, [(folder, destination_folder) for folder in folder_list])

    print("✅ All files copied.")


if __name__ == "__main__":
    source_root = r'C:\Users\isset\Downloads\hansards_2025'
    destination = r'C:\Users\isset\Downloads\destination_5'
    start_time = time.time()
    copy_files_from_deep_folders_parallel(source_root, destination)
    end_time = time.time()
    print(f"\n✅ Finished in {end_time - start_time:.2f} seconds.")

