import os
import multiprocessing
import subprocess
import time


def run_worker_script(source_path, dest_path):
    max_retries = 3
    for attempt in range(max_retries):
        result = subprocess.run(
            ["python", "convert_worker.py", source_path, dest_path],
            capture_output=True,
            text=True
        )
        if result.returncode == 0:
            return
        else:
            print(result.stderr)
            print(f"Retrying {source_path} in 3s... (Attempt {attempt+1}/{max_retries})")
            time.sleep(3)
    print(f"❌ Failed after retries: {os.path.basename(source_path)}")


def convert_all_docs_to_docx_parallel(source_folder, dest_folder):
    os.makedirs(dest_folder, exist_ok=True)
    tasks = []

    for filename in os.listdir(source_folder):
        if filename.endswith(".doc") and not filename.endswith(".docx"):
            source_path = os.path.join(source_folder, filename)
            base_name = os.path.splitext(filename)[0]
            dest_path = os.path.join(dest_folder, base_name + ".docx")
            tasks.append((source_path, dest_path))

    # Use 2–4 workers max for Word COM safety
    with multiprocessing.Pool(processes=6) as pool:
        pool.starmap(run_worker_script, tasks)

def run_worker_script(source_path, dest_path):
    subprocess.run(["python", "convert_worker.py", source_path, dest_path])


if __name__ == "__main__":
    convert_all_docs_to_docx_parallel(
        r'C:\Users\isset\Downloads\destination_1',
        r'C:\Users\isset\Downloads\destination_2'
    )
