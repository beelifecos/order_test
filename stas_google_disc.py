import re
import os
import tempfile
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
    if any(k in name_lower for k in ["선크림", "sun screen", "선 ","sun","선 크림" , "spf", "sun cream","sun stick", "sun care","선스틱"]):
        return "SUN CARE I ЗАЩИТА ОТ СОЛНЦА"
    if any(k in name_lower for k in ["미셀라", "micellar","ENZYME", "필링","Cleanser", "peeling","비누","soup","코팩","nose pack", "클렌징","리무버 ","remover", "cleansing","필 오프 팩 ","peel off pack", "폼", "foam", "필링 젤", "peeling gel", "클렌저 오일", "cleanser oil ", "오일 클렌저", "oil cleanser", "마일드", "mild", "워터","cleansing water", "water wash"]):
        return "CLEANSING I ОЧИЩЕНИЕ"
    if any(k in name_lower for k in ["앰플", "ampoule","진액","스킨","유연액","에멀션","유연수","patch","pad","REEDLE SHOT","pack","source","Moisturizer", "ampule","멀티밤", "multi balm","에멀전", "소프너","softner", "크림", "cream", "토너","아이패치","eye patch","멀티 밤 ", "toner","에멀젼","emulsion","엑스트라 액터","수액","리프샷", "유액", "마스크", "mask", "에센스", "essence","옴므 올인원", "세럼", "serum", "아이크림", "eye cream", "eye serum", "하이드레이팅", "hydrating", "비타", "vitamin", "리프팅", "lifting", "미백", "whitening", "brightening", "수딩", "soothing", "balm", "concentrate","패드","링클 진액고","수분팩","앰풀오일","진액 오일","아이리프트","멀티스틱","밸런서"]):
        return "SKIN CARE I УХОД ЗА ЛИЦОМ"
    if any(k in name_lower for k in ["바디", "body", "로션", "lotion","여성 청결제", "스크럽", "scrub", "바디워시","넥", "body wash", "샤워젤", "shower gel","여성청결제"]):
        return "BODY CARE I УХОД ЗА ТЕЛОМ"
    if any(k in name_lower for k in ["샴푸", "shampoo","왁싱 매니큐어","미쟝센","헤어커버","LPP 트리트","아르드포 스프레이","염색", "컨디셔너","일진 케론 시스테인 웨이브","퍼퓸 린스", "conditioner","아이 팔레트", "린스","트리트먼트","hair treatment", "헤어 린스","쿨링 토닉","케라틴", "헤어칼라","크리닉 칼라"," 헤어 칼라"," 헤어","스타일링 무스 ","셋팅 스프레이", "hair", "treatment", "헤어팩","시스테인","헤어비비", "hair pack","새치", "헤어오일", "hair oil"]):
        return "HAIR CARE I УХОД ЗА ВОЛОСАМИ"
    if any(k in name_lower for k in ["립", "lip", "파운데이션","jelly stick", "foundation", "블러셔", "blush","섀도 팔레트","shedow", "섀도우"," 마스카라 ","mascara", "비비","프라이머","골든 베이스","베이스","bb cream", "아이브로우","eye brow", "팩솔","eye liner","아이라이너","블러쉬","blasher","아이브로우 펜슬","pencil", "물광글로우"," glow" , "컨실러","concealer","펜슬 ","펜 라이너","펜 라이너","liner", "브러쉬 라이너","하이라이터", "hilighter", "쉐도우", "eyeshadow", "글로스", "아이섀도", "투웨이케익", "two way cake", "스킨커버","cover","eye shadow", "메이크업", "make up","팩트","pact","파우더","powder"," 피니쉬","finish", "base","컨투어 "," 미스트", "쿠션", "cushion", "틴트", "tint","베이스 핑크"]):
        return "MAKE UP I ДЕКОРАТИВНЫЙ МАКИЯЖ"
    if any(k in name_lower for k in ["세트", "set", "기획세트","기획", "special set", "패키지", "package", "컬렉션", "collection","3종","kit","키트","세트","기품세트","궁중세트","기획","종세트"]):
        return "SKIN CARE SET I УХОДОВЫЕ НАБОРЫ"
    if any(k in name_lower for k in ["남성", "men","보닌", "스프레이 드라이 임팩트","포맨", " 애프터 쉐이브 ", "for men","쉐이브","homme"]):
        return "FOR MEN / Для мужчин"
    if any(k in name_lower for k in ["샘플", "sample", "미니", "mini", "트래블", "travel"]):
        return "SAMPLE | ПРОБНИКИ"
    if any(k in name_lower for k in ["건강기능식품", "supplement", "비타민", "vitamin", "오메가", "omega", "프로바이오틱스", "probiotic","boto"]):
        return "БАДЫ"
    if any(k in name_lower for k in ["코롱","데오드란트","bag","perfume","코치","부쉐론","메디안","쇼핑백","향수" "폴로","brush","메르세데스 벤츠"," 치약 ","엘리자베스아덴 ","샤워볼","주방세제","세정제","공용기","헤어롤","베르사체","버버리","버블제로","구찌 ","코가위","족집게","오데퍼퓸","쌍꺼풀"," toothpast","화장솜","스프링밴드","4D 페이셜","메디안 "," 뷰티 바","면봉","불가리","손톱전용","물티슈","때비누","몽블랑","롤리타","세탁비누","고무장갑","씨케이","에스티로더","페리오","제습혁명","웰투스","엘지","손소독제","지미추","엘지 테크","네일 스티커","뚜왈렛","씨케이","랑방","폴로","SPPC","습기제거제","각티슈","폴로 스포츠","장아떼","키친타올","2080","위생롤백","모스키노 ","디퓨저","입욕제","겐조","돌체 앤 가바나","아리아나 그란데","퍼퓸","에르메스","샤워코롱","존 바바토스","로페스 매니큐어","매니큐어"]):
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
        # Добавьте остальные бренды
    }
    return brand_name_map.get(brand_cd, brand_cd)

def handle_alert(driver):
    try:
        WebDriverWait(driver, 2).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
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
    # добавьте шаги логина через driver.find_element(...)

    # --- Список брендов ---
    brand_urls = [
        "https://example.com?brand_cd=BR000357",
        "https://example.com?brand_cd=BR001115",
        # добавьте остальные
    ]

    for brand_url in brand_urls:
        brand_name = extract_brand_name(brand_url)
        print(f"Scraping products for brand: {brand_name}")

        driver.get(brand_url)
        handle_alert(driver)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "album")))

        page_links = driver.find_elements(By.CLASS_NAME, "page-link")
        num_pages = len(page_links) if page_links else 1
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
                    if pieces_per_box_element:
                        pieces_per_box = pieces_per_box_element.text.split(':')[-1].strip().replace('ea','').replace(')','').replace(',','')
                        if not pieces_per_box or pieces_per_box == '':
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
                    print("Error parsing product:", e)

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
                print("❌ Ошибка при загрузке файла на Google Drive:", e)

            # --- Переход на следующую страницу ---
            if page_num < num_pages:
                try:
                    next_page_button = driver.find_element(By.XPATH, f"//a[@class='page-link' and @page='{page_num + 1}']")
                    next_page_button.click()
                except Exception as e:
                    print("⚠️ Error clicking next page:", e)
                    break

    driver.quit()
    print("🎯 Scraping completed.")

# --- Запуск парсера ---
if __name__ == "__main__":
    login_and_scrape("beelifecos", "1983beelif")
