import os
import json
from pathlib import Path
from datetime import datetime
from playwright.sync_api import sync_playwright
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

SITE_URL = "https://wholesale.stylekorean.com/"
USERNAME = os.getenv("SITE_USER", "beelifecos")
PASSWORD = os.getenv("SITE_PASS", "1983beelif")
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

def save_json(data):
    ts = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    file_path = OUTPUT_DIR / f"stylekorean_{ts}.json"
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return file_path

def upload_to_drive(file_path):
    # авторизация через client_secrets.json
    gauth = GoogleAuth()
    gauth.LoadClientConfigFile("client_secrets.json")
    gauth.LocalWebserverAuth()  # при Actions можно заменить на gauth.CommandLineAuth()
    drive = GoogleDrive(gauth)
    gfile = drive.CreateFile({'title': file_path.name})
    gfile.SetContentFile(str(file_path))
    gfile.Upload()
    print("Uploaded to Google Drive:", file_path.name)

def parse_products_from_page(page):
    products = []
    cards = page.query_selector_all("div.product, .product-item, li.product, .card")
    for c in cards:
        try:
            title_el = c.query_selector("h2, h3, .name, .product-title, .product-name")
            title = title_el.inner_text().strip() if title_el else None

            price_el = c.query_selector(".price, .product-price")
            price = price_el.inner_text().strip() if price_el else None

            link_el = c.query_selector("a[href]") or c
            href = link_el.get_attribute("href") if link_el else None
            if href and href.startswith("/"):
                href = SITE_URL.rstrip("/") + href

            img_el = c.query_selector("img")
            img = img_el.get_attribute("src") if img_el else None
            if img and img.startswith("/"):
                img = SITE_URL.rstrip("/") + img

            products.append({"title": title, "price": price, "link": href, "image": img})
        except:
            continue
    return products

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
        context = browser.new_context()
        page = context.new_page()
        page.goto(SITE_URL)
        # логин (подредактировать селекторы, если нужно)
        try:
            user = page.query_selector("input[type=email], input[name='email'], input[name='username']")
            pwd = page.query_selector("input[type=password]")
            btn = page.query_selector("button[type=submit]")
            if user and pwd:
                user.fill(USERNAME)
                pwd.fill(PASSWORD)
                btn.click()
                page.wait_for_load_state("networkidle", timeout=15000)
        except:
            pass
        page.goto(SITE_URL + "collections/all", wait_until="networkidle")
        for _ in range(5):
            page.mouse.wheel(0, 1000)
            page.wait_for_timeout(500)
        products = parse_products_from_page(page)
        browser.close()

    file_path = save_json(products)
    print("Saved locally:", file_path)
    upload_to_drive(file_path)

if __name__ == "__main__":
    run()
