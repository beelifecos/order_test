import re
import os
import tempfile
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from bs4 import BeautifulSoup
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# --- Функция для распределения по категориям ---
def assign_category(name):
    if not name:
        return "НЕОПРЕДЕЛЕНО"
    name_lower = name.lower()
    if any(k in name_lower for k in ["선크림", "sun screen", "spf", "sun care"]):
        return "SUN CARE I ЗАЩИТА ОТ СОЛНЦА"
    if any(k in name_lower for k in ["미셀라", "micellar", "cleanser", "peeling"]):
        return "CLEANSING I ОЧИЩЕНИЕ"
    if any(k in name_lower for k in ["앰플", "ampoule","cream","serum","moisturizer"]):
        return "SKIN CARE I УХОД ЗА ЛИЦОМ"
    if any(k in name_lower for k in ["바디", "body", "lotion","scrub","shower"]):
        return "BODY CARE I УХОД ЗА ТЕЛОМ"
    if any(k in name_lower for k in ["샴푸","shampoo","conditioner","hair"]):
        return "HAIR CARE I УХОД ЗА ВОЛОСАМИ"
    if any(k in name_lower for k in ["립", "lip", "foundation","blush","makeup"]):
        return "MAKE UP I ДЕКОРАТИВНЫЙ МАКИЯЖ"
    if any(k in name_lower for k in ["세트", "set", "kit"]):
        return "SKIN CARE SET I УХОДОВЫЕ НАБОРЫ"
    if any(k in name_lower for k in ["남성", "men","for men"]):
        return "FOR MEN / Для мужчин"
    if any(k in name_lower for k in ["샘플", "sample","mini","travel"]):
        return "SAMPLE | ПРОБНИКИ"
    if any(k in name_lower for k in ["supplement", "vitamin","omega","probiotic"]):
        return "БАДЫ"
    if any(k in name_lower for k in ["perfume","bag","toothpaste","hand sanitizer"]):
        return "ТОВАРЫ ДЛЯ ДОМА И ЗДОРОВЬЯ"
    return "НЕОПРЕДЕЛЕНО"

# --- Функции для работы с брендом ---
def extract_brand_name(brand_url):
    brand_cd = brand_url.split("brand_cd=")[-1]
    brand_name_map = {
        "BR000357": "9Wishes",
        "BR001115": "ABEREDE",
        "BR000311": "Abib",
        "BR000067": "ACWELL",
        "BR000473": "AESTURA",
        "BR000457": "AHEADS",
        "BR000091": "A.H.C",
        # добавьте остальные бренды
    }
    return brand_name_map.get(brand_cd, brand_cd)

def handle_alert(driver):
    try:
        WebDriverWait(driver, 2).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
        print("⚠️ Alert accepted")
    except:
        pass

# --- Основной парсер ---
def login_and_scrape(username, password):
    # --- Настройка Google Drive ---
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)
    folder_id = "ВАШ_FOLDER_ID"  # замените на ID папки
    file_name = "products.xlsx"
    file_path = os.path.join(os.getcwd(), file_name)

    # --- Настройка Excel ---
    wb = Workbook()
    ws = wb.active
    ws.append([
        "img_src","brand_name","name","category","unit","moq","quantity",
        "pieces_per_box","item_code","product_code","price_discounted",
        "cena_na_site","price","lang","pieces_per_box2","all","qty","price_old",
        "STATUS","status_value","procent"
    ])

    # --- Настройка Selenium ---
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-notifications")

    tmp_dir = tempfile.mkdtemp()
    options.add_argument(f"--user-data-dir={tmp_dir}")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    # --- Логин ---
    driver.get("URL_ВОЙТИ")  # замените на URL входа
    print("⚡ Открыта страница входа")
    # TODO: добавьте шаги логина через driver.find_element(...)

    # --- Список брендов ---
    brand_urls = [
        "https://example.com?brand_cd=BR000357",
        "https://example.com?brand_cd=BR001115",
        # добавьте остальные
    ]

    for brand_url in brand_urls:
        brand_name = extract_brand_name(brand_url)
        print(f"📦 Scraping products for brand: {brand_name}")

        driver.get(brand_url)
        handle_alert(driver)
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "album")))
        except:
            print(f"❌ Элемент 'album' не найден на странице {brand_url}")
            continue

        page_links = driver.find_elements(By.CLASS_NAME, "page-link")
        num_pages = len(page_links) if page_links else 1
        if len(page_links) >= 3:
            num_pages_element = page_links[-3]
            num_pages_label = num_pages_element.get_attribute("aria-label")
            if num_pages_label:
                try:
                    num_pages = int(num_pages_label.split()[-1])
                except:
                    num_pages = 1

        for page_num in range(1, num_pages + 1):
            print(f"🔹 Обрабатываем страницу {page_num}/{num_pages}")
            handle_alert(driver)
            time.sleep(3)  # небольшая пауза для полной загрузки
            try:
                WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "album")))
            except:
                print(f"❌ Элемент 'album' не найден на странице {page_num}")
                continue

            soup = BeautifulSoup(driver.page_source, 'html.parser')
            product_cards = soup.find_all("div", class_="card mb-4 shadow-sm")

            for card in product_cards:
                try:
                    name = card.find("span", class_="productTxt").text.strip()
                    category = assign_category(name)
                    item_code = card.find("span", class_="productCodeTxt").text.split("SKU:")[1].strip()
                    quantity_availabl_element = card.find("span", class_="qtyTxt")
                    quantity_availabl = quantity_availabl_element.text.strip().replace('ea','').split()[0].replace(',','') if quantity_availabl_element else None
                    img_element = card.find("img", class_="Img_Product")
                    img_src = img_element.get('src') if img_element else None
                    moq_element = card.find("span", class_="moqTxt")
                    moq = moq_element.text.split(":")[-1].replace("ea","").strip() if moq_element else None
                    Product_code = card.find("span", class_="barcodeTxt").text.strip().split(":")[1].strip()
                    pieces_per_box_element = card.find("span", class_="boxCnt")
                    if pieces_per_box_element:
                        pieces_per_box = pieces_per_box_element.text.split(':')[-1].strip().replace('ea','').replace(')','').replace(',','')
                        if not pieces_per_box:
                            pieces_per_box = '20'
                    else:
                        pieces_per_box = '20'
                    price_discounted_element = card.find("span", class_="priceTxt")
                    price_discounted = float(price_discounted_element.text.strip().replace("KRW","").replace(",","").replace(".00","")) if price_discounted_element else 0
                    price_old_element = card.find("span", class_="priceOld2")
                    price_old = float(price_old_element.text.strip().replace("KRW","").replace(",","").replace(".00","")) if price_old_element else None
                    cena_na_site = round(price_discounted * 1.2 / 1250, 2)
                    price = round(price_discounted * 1.1 / 1250, 2)
                    cena_na_site_str = f"{cena_na_site:.2f}".replace(",", ".")
                    price_str = f"{price:.2f}".replace(",", ".")

                    item_code_clean = re.sub(r'\s+', '', item_code)
                    product_code_clean = re.sub(r'\s+', '', Product_code)
                    status_value = f"Бренд///{brand_name[0].upper()}///{brand_name}"
                    STATUS="A"
                    procent= round(price_discounted / price_old , 2) if price_old else 0

                    ws.append([
                        img_src, brand_name, name, category, 'ea', moq, quantity_availabl,
                        pieces_per_box, item_code_clean, product_code_clean, price_discounted,
                        cena_na_site_str, price_str, 'ru', pieces_per_box, 'Все', '1', price_old,
                        STATUS, status_value, procent
                    ])
                except Exception as e:
                    print("❌ Error parsing product:", e)

            # --- Сохраняем Excel локально ---
            try:
                wb.save(file_path)
                print(f"✅ File saved successfully after page {page_num}")
            except Exception as e:
                print("❌ Error saving file:", e)

            # --- Загружаем на Google Drive ---
            try:
                query = f"'{folder_id}' in parents and trashed=false and title='{file_name}'"
                file_list = drive.ListFile({'q': query}).GetList()

                if file_list:
                    file_drive = file_list[0]
                    print(f"Файл найден, обновляем содержимое...")
                else:
                    file_drive = drive.CreateFile({'title': file_name, 'parents': [{'id': folder_id}]})
                    print(f"Файл не найден, создаём новый...")

                file_drive.SetContentFile(file_path)
                file_drive.Upload()
                print(f"✅ Файл успешно обновлён на Google Drive после страницы {page_num}")

            except Exception as e:
                print(f"❌ Ошибка при загрузке файла на Google Drive: {e}")

            # --- Переход на следующую страницу ---
            if page_num < num_pages:
                try:
                    next_page_button = driver.find_element(By.XPATH, f"//a[@class='page-link' and @page='{page_num + 1}']")
                    next_page_button.click()
                except Exception as e:
                    print(f"⚠️ Error clicking next page: {e}")
                    break

    driver.quit()
    print("🎯 Scraping completed.")

# --- Запуск парсера ---
if __name__ == "__main__":
    login_and_scrape("beelifecos", "1983beelif")
