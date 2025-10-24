import re
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
import os

# --- Функция для распределения по категориям ---
def assign_category(name):
    if not name:
        return "НЕОПРЕДЕЛЕНО"
    name_lower = name.lower()
    if any(k in name_lower for k in ["선크림", "sun screen", "spf", "sun cream", "sun stick", "sun care"]):
        return "SUN CARE I ЗАЩИТА ОТ СОЛНЦА"
    if any(k in name_lower for k in ["미셀라", "micellar","cleansing","peeling","foam","oil cleanser","toner","mask","essence","serum","eye cream"]):
        return "SKIN CARE I УХОД ЗА ЛИЦОМ"
    if any(k in name_lower for k in ["바디", "body", "lotion","scrub","body wash"]):
        return "BODY CARE I УХОД ЗА ТЕЛОМ"
    if any(k in name_lower for k in ["샴푸","shampoo","conditioner","hair","treatment","hair pack","hair oil"]):
        return "HAIR CARE I УХОД ЗА ВОЛОСАМИ"
    if any(k in name_lower for k in ["립", "lip","foundation","blush","mascara","bb cream","concealer","tint","cushion"]):
        return "MAKE UP I ДЕКОРАТИВНЫЙ МАКИЯЖ"
    if any(k in name_lower for k in ["세트","set","package","kit","collection"]):
        return "SKIN CARE SET I УХОДОВЫЕ НАБОРЫ"
    if any(k in name_lower for k in ["남성","men","for men","homme"]):
        return "FOR MEN / Для мужчин"
    if any(k in name_lower for k in ["샘플","sample","mini","travel"]):
        return "SAMPLE | ПРОБНИКИ"
    if any(k in name_lower for k in ["supplement","vitamin","omega","probiotic"]):
        return "БАДЫ"
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
        "BR000487": "AIRIVE",
        "BR000811": "AKF",
        "BR000502": "ALETHEIA",
        "BR001097": "ALLIONE",
        "BR000081": "Amos",
        "BR000365": "AMPLE N",
        "BR000572": "AMTS (All My Things)",
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

def handle_alert(driver):
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
    except:
        pass

# --- Основная функция скрапинга ---
def login_and_scrape(username, password):
    options = Options()
    # options.add_argument("--headless=new")  # на время теста можно отключить
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-notifications")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    driver.get("https://wholesale.stylekorean.com/Member/SignIn")
    handle_alert(driver)

    # --- Ждём поле логина ---
    WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.ID, "login_id")))
    driver.find_element(By.ID, "login_id").send_keys(username)
    driver.find_element(By.ID, "pwd").send_keys(password)

    WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, ".Btn_Login[type='submit']"))
    )
    driver.find_element(By.CSS_SELECTOR, ".Btn_Login[type='submit']").click()
    handle_alert(driver)

    # --- Проверка успешного входа ---
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".gnbMy"))
        )
        print("✅ Вход выполнен успешно!")
    except:
        print("❌ Не удалось войти. Проверьте логин/пароль или капчу.")
        driver.quit()
        return

    # --- Excel ---
    wb = Workbook()
    ws = wb.active
    ws.append([
        "Изображение", "Бренд", "Наименование", "Категория", "Единица измерения",
        "MOQ", "Фактический остаток", "in box", "Артикул", "Product code",
        "Цена Discounted KRW", "Cena na site $", "Price", "Language", "Lower limit",
        "User group", "Особенности", "Старая цена KRW","status","category","procent"
    ])

    file_path = 'parser_stas_final_2.xlsx'

    # --- Google Drive ---
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)
    folder_id = "10J-E4RcBJFfrdcqU_UAWask8BKTZ5Mw2"
    file_name = os.path.basename(file_path)

    brand_urls = [
        "https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000357", # 9Wishes
        "https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001115", # ABEREDE
        "https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000311", # Abib
        # ... остальные ссылки
    ]

    for brand_url in brand_urls:
        brand_name = extract_brand_name(brand_url)
        print(f"Scraping products for brand: {brand_name}")

        driver.get(brand_url)
        handle_alert(driver)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "album")))

        # --- Определяем количество страниц ---
        page_links = driver.find_elements(By.CLASS_NAME, "page-link")
        num_pages = 1
        if len(page_links) >= 3:
            num_pages_element = page_links[-3]
            num_pages_label = num_pages_element.get_attribute("aria-label")
            if num_pages_label:
                num_pages = int(num_pages_label.split()[-1])

        for page_num in range(1, num_pages + 1):
            handle_alert(driver)
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "album")))
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
                    pieces_per_box = pieces_per_box_element.text.split(':')[-1].strip().replace('ea','').replace(')','').replace(',','') if pieces_per_box_element else '20'
                    price_discounted_element = card.find("span", class_="priceTxt")
                    price_discounted = float(price_discounted_element.text.strip().replace("KRW","").replace(",","").replace(".00","")) if price_discounted_element else 0
                    price_old_element = card.find("span", class_="priceOld2")
                    price_old = float(price_old_element.text.strip().replace("KRW","").replace(",","").replace(".00","")) if price_old_element else None
                    cena_na_site = round(price_discounted * 1.2 / 1250, 2)
                    price = round(price_discounted * 1.1 / 1250, 2)
                    ws.append([
                        img_src, brand_name, name, category, 'ea', moq, quantity_availabl,
                        pieces_per_box, re.sub(r'\s+','',item_code), re.sub(r'\s+','',Product_code),
                        price_discounted, f"{cena_na_site:.2f}", f"{price:.2f}", 'ru',
                        pieces_per_box, 'Все', '1', price_old, "A",
                        f"Бренд///{brand_name[0].upper()}///{brand_name}",
                        round(price_discounted / price_old,2) if price_old else 0
                    ])
                except Exception as e:
                    print("Error parsing product:", e)

            # --- Сохраняем Excel ---
            try:
                wb.save(file_path)
                print(f"✅ File saved successfully after page {page_num}")
            except Exception as e:
                print("❌ Error saving file:", e)

            # --- Google Drive ---
            try:
                query = f"'{folder_id}' in parents and trashed=false and title='{file_name}'"
                file_list = drive.ListFile({'q': query}).GetList()
                if file_list:
                    file_drive = file_list[0]
                    print(f"Обновляем существующий файл '{file_name}'")
                else:
                    file_drive = drive.CreateFile({'title': file_name, 'parents':[{'id': folder_id}]})
                    print(f"Создаём новый файл '{file_name}'")
                file_drive.SetContentFile(file_path)
                file_drive.Upload()
                print(f"✅ Файл '{file_name}' обновлён на Google Drive")
            except Exception as e:
                print("❌ Ошибка Google Drive:", e)

            # --- Переход на следующую страницу ---
            if page_num < num_pages:
                try:
                    next_page_button = driver.find_element(By.XPATH, f"//a[@class='page-link' and @page='{page_num+1}']")
                    next_page_button.click()
                except Exception as e:
                    print("⚠️ Error clicking next page:", e)
                    break

    driver.quit()
    print("🎯 Scraping completed.")

# --- Запуск ---
login_and_scrape("beelifecos", "1983beelif")
