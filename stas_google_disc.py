import requests
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
import logging
import time

# ------------------------- Настройки -------------------------
USERNAME = "beelifecos"
PASSWORD = "1983beelif"
brand_urls = [
    "https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000357",
    # добавьте другие URL брендов
]
file_name = "stylekorean_products.xlsx"
file_path = f"./{file_name}"
folder_id = "YOUR_GOOGLE_DRIVE_FOLDER_ID"  # замените на свой ID
# -------------------------------------------------------------

# --- Логирование ошибок ---
logging.basicConfig(filename='errors.log', level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# --- Авторизация Google Drive ---
gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)

# --- Excel ---
wb = Workbook()
ws = wb.active
ws.append([
    "img_src", "brand_name", "name", "category", "unit", "moq", "quantity_availabl",
    "pieces_per_box", "item_code", "product_code", "price_discounted",
    "cena_na_site", "price", "lang", "pieces_per_box_2", "all", "one",
    "price_old", "STATUS", "status_value", "procent"
])

# --- Функции ---
def assign_category(name):
    """Простая функция для категории (можно улучшить под свои правила)."""
    if "Cream" in name or "Mask" in name:
        return "Skincare"
    elif "Lip" in name or "Tint" in name:
        return "Makeup"
    else:
        return "Other"

def extract_brand_name(url):
    """Получаем название бренда из URL"""
    return url.split("brand_cd=")[-1]

# --- Сессия для логина ---
with requests.Session() as session:
    try:
        # 1️⃣ Получаем токен с формы входа
        login_page = session.get("https://wholesale.stylekorean.com/Member/SignIn")
        soup = BeautifulSoup(login_page.text, 'html.parser')
        token_input = soup.find("input", {"name": "__RequestVerificationToken"})
        token = token_input['value'] if token_input else ""

        payload = {
            "user_id": USERNAME,
            "pwd": PASSWORD,
            "__RequestVerificationToken": token,
            "prev_page": ""
        }

        login_response = session.post("https://wholesale.stylekorean.com/Member/SignIn", data=payload)
        if "SignIn" not in login_response.url:
            print("✅ Login successful!")
        else:
            print("❌ Login failed!")
            exit(1)
    except Exception as e:
        logging.error(f"Login error: {e}")
        exit(1)

    # --- Парсинг по брендам ---
    for brand_url in brand_urls:
        brand_name = extract_brand_name(brand_url)
        print(f"Scraping products for brand: {brand_name}")
        page = 1

        while True:
            try:
                r = session.get(brand_url + f"&page={page}")
                soup = BeautifulSoup(r.text, 'html.parser')
                product_cards = soup.find_all("div", class_="card mb-4 shadow-sm")

                if not product_cards:
                    break

                for card in product_cards:
                    try:
                        name = card.find("span", class_="productTxt").text.strip()
                        category = assign_category(name)
                        item_code = card.find("span", class_="productCodeTxt").text.split("SKU:")[-1].strip()
                        quantity_availabl_element = card.find("span", class_="qtyTxt")
                        quantity_availabl = quantity_availabl_element.text.strip().replace('ea','').split()[0].replace(',','') if quantity_availabl_element else None
                        img_element = card.find("img", class_="Img_Product")
                        img_src = img_element.get('src') if img_element else None
                        moq_element = card.find("span", class_="moqTxt")
                        moq = moq_element.text.split(":")[-1].replace("ea","").strip() if moq_element else None
                        Product_code = card.find("span", class_="barcodeTxt").text.strip().split(":")[-1].strip()
                        pieces_per_box_element = card.find("span", class_="boxCnt")
                        if pieces_per_box_element:
                            pieces_per_box = pieces_per_box_element.text.split(':')[-1].strip().replace('ea','').replace(')','').replace(',','')
                            if not pieces_per_box:
                                pieces_per_box = '20'
                        else:
                            pieces_per_box = '20'
                        price_discounted_element = card.find("span", class_="priceTxt")
                        price_discounted = price_discounted_element.text.strip().replace("KRW","").replace(",","").replace(".00","") if price_discounted_element else 0
                        price_discounted = float(price_discounted)
                        price_old_element = card.find("span", class_="priceOld2")
                        price_old = price_old_element.text.strip().replace("KRW","").replace(",","").replace(".00","") if price_old_element else None
                        if price_old:
                            price_old = float(price_old)
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
                        logging.error(f"Error parsing product: {e}")

                # Проверяем наличие кнопки "следующая страница"
                next_button = soup.find("a", class_="page-link", text=str(page + 1))
                if next_button:
                    page += 1
                else:
                    break

                # Сохраняем после каждой страницы
                wb.save(file_path)
                print(f"✅ Saved page {page} to Excel")
                time.sleep(1)  # небольшой таймаут между страницами
            except Exception as e:
                logging.error(f"Error scraping page {page} of {brand_name}: {e}")
                break

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
            print(f"✅ Файл '{file_name}' успешно обновлён на Google Drive")
        except Exception as e:
            logging.error(f"Ошибка при загрузке файла на Google Drive: {e}")

print("🎯 Scraping completed.")
