from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import requests
import os
import time

BASE_URL = "https://parliament.go.ug"
HANARDS_URL = f"{BASE_URL}/hansards"
DOWNLOAD_DIR = "./hansards"

os.makedirs(DOWNLOAD_DIR, exist_ok=True)

def get_soup(url, driver):
    driver.get(url)
    time.sleep(3)
    return BeautifulSoup(driver.page_source, "html.parser")

def scrape_hansards():
    options = Options()
    options.headless = True
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    print(f"Visiting {HANARDS_URL}...")
    soup = get_soup(HANARDS_URL, driver)

    # Step 1: Find all folder links
    folder_links = set()
    for link in soup.find_all("a", href=True):
        href = link['href']
        if "hansards" in href and not href.endswith('.pdf'):
            full_url = BASE_URL + href if href.startswith("/") else href
            folder_links.add(full_url)

    # Step 2: Visit each folder and download PDFs
    for folder_url in folder_links:
        print(f"\nVisiting folder: {folder_url}")
        folder_soup = get_soup(folder_url, driver)

        for link in folder_soup.find_all("a", href=True):
            href = link['href']
            if href.endswith(".pdf"):
                file_url = BASE_URL + href if href.startswith("/") else href
                filename = file_url.split("/")[-1]
                save_path = os.path.join(DOWNLOAD_DIR, filename)

                if not os.path.exists(save_path):
                    print(f"Downloading {filename}...")
                    try:
                        file_data = requests.get(file_url).content
                        with open(save_path, 'wb') as f:
                            f.write(file_data)
                    except Exception as e:
                        print(f"Failed to download {file_url}: {e}")
                else:
                    print(f"Already downloaded: {filename}")

    driver.quit()

if __name__ == "__main__":
    scrape_hansards()
