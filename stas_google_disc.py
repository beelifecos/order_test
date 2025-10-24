import re
import os
import tempfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from bs4 import BeautifulSoup
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# --- Категоризация ---
def assign_category(name):
    if not name:
        return "НЕОПРЕДЕЛЕНО"
    name_lower = name.lower()
    if any(k in name_lower for k in ["선크림", "sun screen", "spf"]):
        return "SUN CARE I ЗАЩИТА ОТ СОЛНЦА"
    if any(k in name_lower for k in ["클렌징", "cleanser", "foam", "peeling"]):
        return "CLEANSING I ОЧИЩЕНИЕ"
    if any(k in name_lower for k in ["앰플", "ampoule", "cream", "serum", "moisturizer"]):
        return "SKIN CARE I УХОД ЗА ЛИЦОМ"
    if any(k in name_lower for k in ["body", "로션", "scrub"]):
        return "BODY CARE I УХОД ЗА ТЕЛОМ"
    if any(k in name_lower for k in ["shampoo", "conditioner", "hair"]):
        return "HAIR CARE I УХОД ЗА ВОЛОСАМИ"
    if any(k in name_lower for k in ["lip", "foundation", "make up", "powder"]):
        return "MAKE UP I ДЕКОРАТИВНЫЙ МАКИЯЖ"
    if any(k in name_lower for k in ["set", "kit", "collection"]):
        return "SKIN CARE SET I УХОДОВЫЕ НАБОРЫ"
    if any(k in name_lower for k in ["men", "for men", "homme"]):
        return "FOR MEN / Для мужчин"
    if any(k in name_lower for k in ["sample", "mini", "travel"]):
        return "SAMPLE | ПРОБНИКИ"
    if any(k in name_lower for k in ["supplement", "vitamin", "omega", "probiotic"]):
        return "БАДЫ"
    if any(k in name_lower for k in ["perfume", "toothpast", "shower"]):
        return "ТОВАРЫ ДЛЯ ДОМА И ЗДОРОВЬЯ"
    return "НЕОПРЕДЕЛЕНО"

# --- Названия брендов ---
def extract_brand_name(brand_url):
    brand_cd = brand_url.split("brand_cd=")[-1]
    brand_name_map = {
        "BR000357": "9Wishes",
        "BR001115": "ABEREDE",
        "BR000311": "Abib",
        "BR000067": "ACWELL",
        "BR000473": "AESTURA",
        "BR000457": "AHEADS",
        "BR000487": "AIRIVE",
        "BR000811": "AKF",
        "BR000502": "ALETHEIA",
        "BR001097": "ALLIONE",
        "BR000081": "Amos",
        "BR000365": "AMPLE N",
        "BR000572": "AMTS",
        "BR000659": "AMUSE",
        "BR000563": "And:ar",
        "BR000522": "ANN 365",
        "BR000516": "ANUA",
        "BR000181": "Apieu",
        "BR001129": "APLB",
        "BR000152": "APRIL SKIN",
        "BR000294": "aromatica",
        "BR000625": "ATHINGS",
        "BR000367": "ATOPALM",
        "BR000558": "ATVT",
        "BR000301": "Avajar",
        "BR000537": "AXIS-Y"
    }
    return brand_name_map.get(brand_cd, "Unknown Brand")

# --- Обработка alert ---
def handle_alert(driver, timeout=5):
    try:
        WebDriverWait(driver, timeout).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
        print("✅ Alert accepted")
    except:
        pass

# --- Основной скрапинг ---
def login_and_scrape(username, password):
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-notifications")
    options.add_argument(f"--user-data-dir={tempfile.mkdtemp()}")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    # --- Логин ---
    driver.get("https://wholesale.stylekorean.com/Member/SignIn")
    handle_alert(driver)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "user_id")))
    driver.find_element(By.ID, "user_id").send_keys(username)
    driver.find_element(By.ID, "pwd").send_keys(password)
    driver.find_element(By.CSS_SELECTOR, ".Btn_Login[type='submit']").click()
    handle_alert(driver)

    # --- Excel ---
    wb = Workbook()
    ws = wb.active
    ws.append([
        "Изображение", "Бренд", "Наименование", "Категория", "Единица измерения",
        "MOQ", "Фактический остаток", "in box", "Артикул", "Product code",
        "Цена Discounted KRW", "Cena na site $", "Price", "Language", "Lower limit",
        "User group", "Особенности", "Старая цена KRW","status","category","procent"
    ])
    file_path = '/Users/tyantamara/parser_stas_final_1.xlsx'

    # --- Google Drive ---
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)
    folder_id = "10J-E4RcBJFfrdcqU_UAWask8BKTZ5Mw2"
    file_name = os.path.basename(file_path)

    brand_urls = [
        "https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000357",
        "https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001115",
        "https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000311",
        "https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000067"
    ]

    for brand_url in brand_urls:
        brand_name = extract_brand_name(brand_url)
        print(f"Scraping products for brand: {brand_name}")

        driver.get(brand_url)
        handle_alert(driver)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "album")))

        # --- Определяем количество страниц ---
        try:
            page_links = driver.find_elements(By.CLASS_NAME, "page-link")
            num_pages = int(page_links[-3].get_attribute("aria-label").split()[-1]) if len(page_links) >= 3 else 1
        except:
            num_pages = 1

        for page_num in range(1, num_pages + 1):
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "album")))
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            product_cards = soup.find_all("div", class_="card mb-4 shadow-sm")

            for card in product_cards:
                try:
                    name = card.find("span", class_="productTxt").text.strip()
                    category = assign_category(name)
                    item_code = card.find("span", class_="productCodeTxt").text.split("SKU:")[-1].strip()
                    quantity = card.find("span", class_="qtyTxt")
                    quantity = quantity.text.strip().replace('ea','').split()[0].replace(',','') if quantity else None
                    img_src = card.find("img", class_="Img_Product").get("src", None)
                    moq = card.find("span", class_="moqTxt")
                    moq = moq.text.split(":")[-1].replace("ea","").strip() if moq else None
                    product_code = card.find("span", class_="barcodeTxt").text.split(":")[-1].strip()
                    pieces_per_box = card.find("span", class_="boxCnt")
                    pieces_per_box = pieces_per_box.text.split(":")[-1].strip().replace('ea','').replace(')','').replace(',','') if pieces_per_box else '20'
                    price_discounted = card.find("span", class_="priceTxt")
                    price_discounted = float(price_discounted.text.strip().replace("KRW","").replace(",","").replace(".00","")) if price_discounted else 0
                    price_old = card.find("span", class_="priceOld2")
                    price_old = float(price_old.text.strip().replace("KRW","").replace(",","").replace(".00","")) if price_old else None
                    cena_na_site = round(price_discounted * 1.2 / 1250, 2)
                    price = round(price_discounted * 1.1 / 1250, 2)

                    ws.append([
                        img_src, brand_name, name, category, 'ea', moq, quantity,
                        pieces_per_box, item_code, product_code, price_discounted,
                        f"{cena_na_site:.2f}", f"{price:.2f}", 'ru', pieces_per_box, 'Все', '1', price_old,
                        "A", f"Бренд///{brand_name[0].upper()}///{brand_name}", round(price_discounted/price_old,2) if price_old else 0
                    ])
                except Exception as e:
                    print(f"⚠️ Error parsing product: {e}")

            # --- Сохраняем Excel ---
            try:
                wb.save(file_path)
                print(f"✅ File saved after page {page_num}")
            except Exception as e:
                print(f"❌ Error saving file: {e}")

            # --- Обновляем Google Drive ---
            try:
                query = f"'{folder_id}' in parents and trashed=false and title='{file_name}'"
                file_list = drive.ListFile({'q': query}).GetList()
                file_drive = file_list[0] if file_list else drive.CreateFile({'title': file_name, 'parents': [{'id': folder_id}]})
                file_drive.SetContentFile(file_path)
                file_drive.Upload()
                print(f"✅ File '{file_name}' uploaded to Google Drive after page {page_num}")
            except Exception as e:
                print(f"❌ Error uploading file to Drive: {e}")

            # --- Переход на следующую страницу ---
            if page_num < num_pages:
                try:
                    next_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, f"//a[@class='page-link' and @page='{page_num + 1}']"))
                    )
                    next_button.click()
                except Exception as e:
                    print(f"⚠️ Error clicking next page: {e}")
                    break

    driver.quit()
    print("🎯 Scraping completed.")

# --- Запуск ---
login_and_scrape("beelifecos", "1983beelif")
