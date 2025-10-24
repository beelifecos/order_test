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
from bs4 import BeautifulSoup
from openpyxl import Workbook
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# -------------------------- Категории --------------------------
def assign_category(name):
    if not name:
        return "НЕОПРЕДЕЛЕНО"
    name_lower = name.lower()
    if any(k in name_lower for k in ["선크림", "sun screen", "sun care","spf","sun stick"]):
        return "SUN CARE I ЗАЩИТА ОТ СОЛНЦА"
    if any(k in name_lower for k in ["미셀라", "micellar","cleansing","폼","foam"]):
        return "CLEANSING I ОЧИЩЕНИЕ"
    if any(k in name_lower for k in ["앰플","ampoule","크림","cream","토너","toner","세럼","serum"]):
        return "SKIN CARE I УХОД ЗА ЛИЦОМ"
    if any(k in name_lower for k in ["바디","body","로션","lotion","scrub","바디워시"]):
        return "BODY CARE I УХОД ЗА ТЕЛОМ"
    if any(k in name_lower for k in ["샴푸","shampoo","컨디셔너","conditioner","hair"]):
        return "HAIR CARE I УХОД ЗА ВОЛОСАМИ"
    if any(k in name_lower for k in ["립","lip","foundation","blush","mascara","concealer"]):
        return "MAKE UP I ДЕКОРАТИВНЫЙ МАКИЯЖ"
    if any(k in name_lower for k in ["세트","set","kit","collection"]):
        return "SKIN CARE SET I УХОДОВЫЕ НАБОРЫ"
    if any(k in name_lower for k in ["남성","men","for men","homme"]):
        return "FOR MEN / Для мужчин"
    if any(k in name_lower for k in ["샘플","sample","mini","travel"]):
        return "SAMPLE | ПРОБНИКИ"
    if any(k in name_lower for k in ["건강기능식품","supplement","vitamin","omega","probiotic"]):
        return "БАДЫ"
    return "НЕОПРЕДЕЛЕНО"

# -------------------------- Бренды --------------------------
def extract_brand_name(brand_url):
    brand_cd = brand_url.split("brand_cd=")[-1]
    brand_name_map = {
        "BR000357": "9Wishes", "BR001115": "ABEREDE", "BR000311": "Abib",
        "BR000067": "ACWELL", "BR000473": "AESTURA", "BR000457": "AHEADS",
        "BR000487": "AIRIVE", "BR000811": "AKF", "BR000502": "ALETHEIA",
        "BR001097": "ALLIONE", "BR000081": "Amos", "BR000365": "AMPLE N",
        "BR000572": "AMTS (All My Things)", "BR000659": "AMUSE", "BR000563": "And:ar",
        "BR000522": "ANN 365", "BR000516": "ANUA", "BR000181": "Apieu",
        "BR001129": "APLB", "BR000152": "APRIL SKIN", "BR000294": "aromatica",
        "BR000625": "ATHINGS", "BR000367": "ATOPALM", "BR000558": "ATVT",
        "BR000301": "Avajar", "BR000537": "AXIS-Y"
    }
    return brand_name_map.get(brand_cd, "Unknown Brand")

# -------------------------- Обработка alert --------------------------
def handle_alert(driver):
    try:
        WebDriverWait(driver, 3).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
    except:
        pass

# -------------------------- Логин --------------------------
def login(driver, username, password):
    driver.get("https://wholesale.stylekorean.com/Member/SignIn")

    WebDriverWait(driver, 15).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.ID, "user_id")))
    WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.ID, "pwd")))

    driver.find_element(By.ID, "user_id").send_keys(username)
    driver.find_element(By.ID, "pwd").send_keys(password)
    driver.find_element(By.CSS_SELECTOR, ".Btn_Login").click()

    # Ждём загрузки страницы после логина
    WebDriverWait(driver, 15).until(
        lambda d: "SignIn" not in d.current_url
    )
    print("✅ Login successful")

# -------------------------- Основной парсер --------------------------
def login_and_scrape(username, password):
    # Настройка headless Chrome
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument(f"--user-data-dir={tempfile.mkdtemp()}")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    login(driver, username, password)

    # Excel
    wb = Workbook()
    ws = wb.active
    ws.append([
        "Изображение", "Бренд", "Наименование", "Категория", "Единица измерения",
        "MOQ", "Фактический остаток", "in box", "Артикул", "Product code",
        "Цена Discounted KRW", "Cena na site $", "Price", "Language", "Lower limit",
        "User group", "Особенности", "Старая цена KRW","status","category","procent"
    ])
    file_path = os.path.join(os.getcwd(), "parser_stas_final.xlsx")

    # Google Drive
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)
    folder_id = "10J-E4RcBJFfrdcqU_UAWask8BKTZ5Mw2"
    file_name = os.path.basename(file_path)

    # Список брендов
    brand_urls = [
        "https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000357",
        "https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001115",
        "https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000311",
        # Добавьте остальные бренды по аналогии
    ]

    for brand_url in brand_urls:
        brand_name = extract_brand_name(brand_url)
        print(f"🔹 Scraping brand: {brand_name}")

        driver.get(brand_url)
        handle_alert(driver)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "album")))

        # Пагинация
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        page_links = driver.find_all(By.CLASS_NAME, "page-link")
        num_pages = 1
        try:
            num_pages = max(int(a.get_attribute("aria-label").split()[-1]) for a in driver.find_elements(By.CLASS_NAME, "page-link") if a.get_attribute("aria-label"))
        except:
            pass

        for page_num in range(1, num_pages + 1):
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "album")))
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            product_cards = soup.find_all("div", class_="card mb-4 shadow-sm")

            for card in product_cards:
                try:
                    name = card.find("span", class_="productTxt").text.strip()
                    category = assign_category(name)
                    item_code = card.find("span", class_="productCodeTxt").text.split("SKU:")[1].strip()
                    quantity_availabl = card.find("span", class_="qtyTxt").text.strip().replace('ea','').split()[0].replace(',','') if card.find("span", class_="qtyTxt") else None
                    img_src = card.find("img", class_="Img_Product")['src'] if card.find("img", class_="Img_Product") else None
                    moq = card.find("span", class_="moqTxt").text.split(":")[-1].replace("ea","").strip() if card.find("span", class_="moqTxt") else None
                    Product_code = card.find("span", class_="barcodeTxt").text.strip().split(":")[1].strip() if card.find("span", class_="barcodeTxt") else None
                    pieces_per_box = card.find("span", class_="boxCnt").text.split(':')[-1].strip().replace('ea','').replace(')','').replace(',','') if card.find("span", class_="boxCnt") else '20'

                    price_discounted = float(card.find("span", class_="priceTxt").text.strip().replace("KRW","").replace(",","").replace(".00","")) if card.find("span", class_="priceTxt") else 0
                    price_old = float(card.find("span", class_="priceOld2").text.strip().replace("KRW","").replace(",","").replace(".00","")) if card.find("span", class_="priceOld2") else None
                    cena_na_site = round(price_discounted * 1.2 / 1250, 2)
                    price = round(price_discounted * 1.1 / 1250, 2)

                    ws.append([
                        img_src, brand_name, name, category, 'ea', moq, quantity_availabl,
                        pieces_per_box, re.sub(r'\s+', '', item_code), re.sub(r'\s+', '', Product_code),
                        price_discounted, f"{cena_na_site:.2f}", f"{price:.2f}", 'ru', pieces_per_box,
                        'Все', '1', price_old, "A", f"Бренд///{brand_name[0].upper()}///{brand_name}", round(price_discounted / price_old, 2) if price_old else 0
                    ])
                except Exception as e:
                    print("❌ Error parsing product:", e)

            # Сохраняем локально
            try:
                wb.save(file_path)
                print(f"✅ File saved successfully after page {page_num}")
            except Exception as e:
                print("❌ Error saving file:", e)

            # Загрузка на Google Drive
            try:
                query = f"'{folder_id}' in parents and trashed=false and title='{file_name}'"
                file_list = drive.ListFile({'q': query}).GetList()

                if file_list:
                    file_drive = file_list[0]
                    print(f"🔹 Updating existing file on Drive: {file_name}")
                else:
                    file_drive = drive.CreateFile({'title': file_name, 'parents': [{'id': folder_id}]})
                    print(f"🔹 Creating new file on Drive: {file_name}")

                file_drive.SetContentFile(file_path)
                file_drive.Upload()
                print(f"✅ File uploaded to Google Drive after page {page_num}")
            except Exception as e:
                print("❌ Google Drive upload error:", e)

            # Переход на следующую страницу
            if page_num < num_pages:
                try:
                    next_page_button = driver.find_element(By.XPATH, f"//a[@class='page-link' and @page='{page_num + 1}']")
                    next_page_button.click()
                except Exception as e:
                    print("⚠️ Error clicking next page:", e)
                    break

    driver.quit()
    print("🎯 Scraping completed.")

# -------------------------- Запуск --------------------------
if __name__ == "__main__":
    login_and_scrape("beelifecos", "1983beelif")
