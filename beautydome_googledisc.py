import time
import re
import urllib.parse
import os
import pickle
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from bs4 import BeautifulSoup

from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from deep_translator import GoogleTranslator


# ---------------- Google Drive -----------------
SCOPES = ['https://www.googleapis.com/auth/drive.file']

def drive_service():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secrets.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('drive', 'v3', credentials=creds)
    return service

def upload_to_drive(file_path, folder_id=None):
    service = drive_service()
    file_metadata = {'name': os.path.basename(file_path)}
    if folder_id:
        file_metadata['parents'] = [folder_id]
    media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
    print(f"Файл загружен на Google Drive. ID: {file['id']}")
    print(f"Ссылка на просмотр: {file['webViewLink']}")
    return file['webViewLink']

# ---------------- Категории товаров -----------------
def assign_category(name):
    if not name:
        return "НЕОПРЕДЕЛЕНО"
    name_lower = name.lower()
    if any(k in name_lower for k in ["선크림", "sun screen", "선 ","sun","선 크림" , "spf", "sun cream","sun stick", "sun care","선스틱"]):
        return "SUN CARE I ЗАЩИТА ОТ СОЛНЦА"
    if any(k in name_lower for k in ["미셀라", "micellar", "필링","cleanser", "peeling","비누","soup","코팩","nose pack", "클렌징","리무버 ","remover", "cleansing","필 오프 팩 ","peel off pack", "폼", "foam", "필링 젤", "peeling gel", "클렌저 오일", "cleanser oil ", "오일 클렌저", "oil cleanser", "마일드", "mild", "워터","cleansing water", "water wash"]):
        return "CLEANSING I ОЧИЩЕНИЕ"
    if any(k in name_lower for k in ["앰플", "ampoule","스킨","유연액","에멀션","유연수","patch","pad","reedle shot","pack","source","moisturizer", "ampule","멀티밤", "multi balm","에멀전", "소프너","softner", "크림", "cream", "토너","아이패치","eye patch","멀티 밤 ", "toner","에멀젼","emulsion","엑스트라 액터","수액","리프샷", "유액", "마스크", "mask", "에센스", "essence","옴므 올인원", "세럼", "serum", "아이크림", "eye cream", "eye serum", "하이드레이팅", "hydrating", "비타", "vitamin", "리프팅", "lifting", "미백", "whitening", "brightening", "수딩", "soothing", "balm", "concentrate","패드","링클 진액고","수분팩","앰풀오일","진액 오일","아이리프트","멀티스틱","밸런서"]):
        return "SKIN CARE I УХОД ЗА ЛИЦОМ"
    if any(k in name_lower for k in ["바디", "body", "로션", "lotion","여성 청결제", "스크럽", "scrub", "바디워시","넥", "body wash", "샤워젤", "shower gel","여성청결제"]):
        return "BODY CARE I УХОД ЗА ТЕЛОМ"
    if any(k in name_lower for k in ["샴푸", "shampoo","왁싱 매니큐어","미쟝센","헤어커버","lpp 트리트","아르드포 스프레이","염색", "컨디셔너","일진 케론 시스테인 웨이브","퍼퓸 린스", "conditioner","아이 팔레트", "린스","트리트먼트","hair treatment", "헤어 린스","쿨링 토닉","케라틴", "헤어칼라","크리닉 칼라"," 헤어 칼라"," 헤어","스타일링 무스 ","셋팅 스프레이", "hair", "treatment", "헤어팩","시스테인","헤어비비", "hair pack","새치", "헤어오일","hair oil"]):
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
def translate_name_to_en(korean_name):
    try:
        if not korean_name or korean_name.strip() == "":
            return ""
        # Переводим с корейского на английский
        translated = GoogleTranslator(source='ko', target='en').translate(korean_name)
        # Немного чистим результат
        cleaned = translated.strip().capitalize()
        return cleaned
    except Exception as e:
        print(f"Ошибка перевода '{korean_name}': {e}")
        return ""

# ---------------- Бренды -----------------
brand_name_map = {
    "1703": {"ko": "가인비책", "en": "GAINBICHAEK"},
    "1700": {"ko": "더후(더히스토리오브후)", "en": "THE HISTORY OF WHOO"},
    "1250": {"ko": "과일나라", "en": "FRUIT NARA"},
    "1252": {"ko": "꽃을든남자", "en": "FLOWER MAN"},
    "1668": {"ko": "끌레드벨", "en": "CLE DE BELLE"},
    "1253": {"ko": "나드리 이노벨라", "en": "NADRI INNOVELLA"},
    "1256": {"ko": "뉴트로지나", "en": "NEUTROGENA"},
    "1257": {"ko": "다나한", "en": "DANAHAN"},
    "1671": {"ko": "도루코", "en": "DORCO"},
    "1689": {"ko": "동성제약", "en": "DONGSUNG PHARM"},
    "1498": {"ko": "드봉", "en": "DEBON"},
    "1268": {"ko": "라미", "en": "RAMI"},
    "1273": {"ko": "루나리스", "en": "LUNARIS"},
    "1424": {"ko": "리엔", "en": "RIEN"},
    "1462": {"ko": "릴랙시아", "en": "RELAXIA"},
    "1280": {"ko": "마몽드", "en": "MAMONDE"},
    "1283": {"ko": "멘소래담", "en": "MENTHOLATUM"},
    "1284": {"ko": "무궁화", "en": "MUGUNGHWA"},
    "1442": {"ko": "미쟝센", "en": "MISE EN SCENE"},
    "1289": {"ko": "바세린", "en": "VASELINE"},
    "1290": {"ko": "바찌", "en": "BAZZI"},
    "1291": {"ko": "백옥생", "en": "BAEKOKSAENG"},
    "1713": {"ko": "베르가모", "en": "BERGAMO"},
    "1736": {"ko": "브에노", "en": "BUENO"},
    "1741": {"ko": "비노아", "en": "VINNOA"},
    "1292": {"ko": "비러브", "en": "BELUV"},
    "1299": {"ko": "산수유", "en": "SANSUYU"},
    "1747": {"ko": "설려", "en": "SEOLLYO"},
    "1494": {"ko": "소망기타", "en": "SOMANG"},
    "1684": {"ko": "숨37도", "en": "SUM37"},
    "1489": {"ko": "쉬림", "en": "SHRIM"},
    "1664": {"ko": "쉬크", "en": "SCHICK"},
    "1702": {"ko": "썬월드", "en": "SUNWORLD"},
    "1734": {"ko": "씨드앤팜", "en": "SEED&PHARM"},
    "1303": {"ko": "아방가드로", "en": "AVANGARDO"},
    "1727": {"ko": "아이차밍", "en": "ICHARMING"},
    "1701": {"ko": "아트피아", "en": "ARTPIA"},
    "1706": {"ko": "알프레도 휘마스", "en": "ALFREDO HUIMAS"},
    "1496": {"ko": "애경", "en": "AEKYUNG"},
    "1676": {"ko": "에띠앙", "en": "ETTIANG"},
    "1308": {"ko": "에바스", "en": "EBAS"},
    "1309": {"ko": "에뿌", "en": "EPPU"},
    "1312": {"ko": "에스클라", "en": "ESCLA"},
    "1678": {"ko": "에스클레어", "en": "ESCLAIR"},
    "1313": {"ko": "에이쓰리에프온", "en": "A3FON"},
    "1315": {"ko": "에코퓨어", "en": "ECO PURE"},
    "1316": {"ko": "엔프라니", "en": "ENPRANI"},
    "1317": {"ko": "엘라스틴", "en": "ELASTINE"},
    "1729": {"ko": "엘지생활건강", "en": "LG HOUSEHOLD & HEALTH"},
    "1711": {"ko": "예지후", "en": "YEJIHU"},
    "1737": {"ko": "예향", "en": "YEHYANG"},
    "1732": {"ko": "오가니아", "en": "OGANIA"},
    "1318": {"ko": "오딧세이", "en": "ODYSSEY"},
    "1714": {"ko": "오릭스", "en": "ORIX"},
    "1320": {"ko": "오퍼스", "en": "OPUS"},
    "1673": {"ko": "오휘(O HUI)", "en": "O HUI"},
    "1321": {"ko": "온더바디", "en": "ON THE BODY"},
    "1322": {"ko": "우드버리", "en": "WOODBURY"},
    "1726": {"ko": "이노벨라", "en": "INNOVELLA"},
    "1440": {"ko": "존슨앤존슨", "en": "JOHNSON & JOHNSON"},
    "1421": {"ko": "쥬리아", "en": "JULIA"},
    "1327": {"ko": "지오", "en": "GIO"},
    "1712": {"ko": "카라코사", "en": "KARAKOSA"},
    "1337": {"ko": "터치미", "en": "TOUCH ME"},
    "1725": {"ko": "팜스테이(명인화장품)", "en": "FARM STAY"},
    "1748": {"ko": "포더스킨", "en": "FOR THE SKIN"},
    "1345": {"ko": "푸드어홀릭", "en": "FOODAHOLIC"},
    "1423": {"ko": "프린시아", "en": "PRINCIA"},
    "1347": {"ko": "피어리스", "en": "PEERLESS"},
    "1349": {"ko": "한불", "en": "HANBUL"},
    "1491": {"ko": "해피바스", "en": "HAPPY BATH"},
    "1495": {"ko": "황후빈", "en": "HWANGHUBIN"}
}

def extract_brand_name(brand_url):
    query = urllib.parse.urlparse(brand_url).query
    params = urllib.parse.parse_qs(query)
    brand_cd = params.get("cno1", [""])[0]
    brand_info = brand_name_map.get(brand_cd, {"ko": "Unknown Brand", "en": "Unknown Brand"})
    return brand_info["ko"], brand_info["en"]

def brand_special_column(brand_name_ko, brand_name_en):
    if not brand_name_en or brand_name_en.strip() == "":
        return f"{brand_name_ko}///X///UNKNOWN"
    first_letter = brand_name_en.strip()[0].upper()
    return f"Бренд///{first_letter}///{brand_name_en.strip().upper()}"

def handle_alert(driver):
    try:
        WebDriverWait(driver, 1).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
    except:
        pass

def add_page_to_url(url, page_num):
    if "page=" in url:
        return re.sub(r'page=\d+', f'page={page_num}', url)
    separator = '&' if '?' in url else '?'
    return f"{url}{separator}page={page_num}"

# ---------------- Scraper -----------------
def login_and_scrape(username, password):
    options = Options()
    options.add_argument('--disable-notifications')
    options.add_argument("start-maximized")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36"
    )
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    # Вход
    driver.get("https://www.beautydome.co.kr/member/login.php")
    handle_alert(driver)
    driver.find_element(By.ID, "login_id").send_keys(username)
    driver.find_element(By.ID, "login_pwd").send_keys(password)
    driver.find_element(By.CSS_SELECTOR, ".box_btn.circle input[type='submit']").click()
    handle_alert(driver)

    # Excel
    wb = Workbook()
    ws = wb.active
    ws.append([
        "Изображение", "Бренд", "Название", "name_en","Единица измерения", "MOQ", "Остаток",
        "Цена", "Цена розницы", "Артикул", "Excel формула", "Категория",
        "Language", "Lower limit", "User group", "Особенности","brand_name_en", "price","cena_na_site","Status",
    ])

    seen_items = set()

    brand_urls = [
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1703", # 가인비책
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1700", # 더후
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1250", # 과일나라
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1252", # 꽃을든남자
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1668", # 끌레드벨
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1253", # 나드리 이노벨라
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1256", # 뉴트로지나
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1257", # 다나한
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1671", # 도루코
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1689", # 동성제약
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1498", # 드봉
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1268", # 라미
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1273", # 루나리스
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1424", # 리엔
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1462", # 릴랙시아
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1280", # 마몽드
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1283", # 멘소래담
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1284", # 무궁화
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1442", # 미쟝센
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1289", # 바세린
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1290", # 바찌
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1291", # 백옥생
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1713", # 베르가모
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1736", # 브에노
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1741", # 비노아
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1292", # 비러브
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1299", # 산수유
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1747", # 설려
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1494", # 소망기타
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1684", # 숨37도
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1489", # 쉬림
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1664", # 쉬크
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1702", # 썬월드
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1734", # 씨드앤팜
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1303", # 아방가드로
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1727", # 아이차밍
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1701", # 아트피아
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1706", # 알프레도 휘마스
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1496", # 애경
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1676", # 에띠앙
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1308", # 에바스
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1309", # 에뿌
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1312", # 에스클라
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1678", # 에스클레어
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1313", # 에이쓰리에프온
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1315", # 에코퓨어
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1316", # 엔프라니
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1317", # 엘라스틴
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1729", # 엘지생활건강
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1711", # 예지후
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1737", # 예향
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1732", # 오가니아
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1318", # 오딧세이
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1714", # 오릭스
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1320", # 오퍼스
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1673", # 오휘(O HUI)
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1321", # 온더바디
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1322", # 우드버리
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1726", # 이노벨라
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1440", # 존슨앤존슨
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1421", # 쥬리아
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1327", # 지오
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1712", # 카라코사
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1337", # 터치미
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1725", # 팜스테이(명인화장품)
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1748", # 포더스킨
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1345", # 푸드어홀릭
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1423", # 프린시아
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1347", # 피어리스
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1349", # 한불
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1491", # 해피바스
        "https://www.beautydome.co.kr/shop/big_section.php?cno1=1495", # 황후빈
    ]

    for brand_url in brand_urls:
        brand_name_ko, brand_name_en = extract_brand_name(brand_url)
        brand_column = brand_special_column(brand_name_ko, brand_name_en)
        print(f"Scraping products for brand: {brand_column}")

        for page_num in range(1, 11):
            page_url = add_page_to_url(brand_url, page_num)
            driver.get(page_url)
            time.sleep(3)

            try:
                WebDriverWait(driver, 1).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.info"))
                )
            except:
                print(f"Страница {page_num} бренда {brand_column} пустая. Прерываем цикл.")
                break

            soup = BeautifulSoup(driver.page_source, 'html.parser')
            products = soup.select("div.info")
            if not products:
                print(f"Страница {page_num} бренда {brand_column} пустая. Прерываем цикл.")
                break

            print(f"Страница {page_num} бренда {brand_column} - найдено {len(products)} товаров")

            for product in products:
                try:
                    name_tag = product.select_one("p.name a")
                    if not name_tag:
                        continue
                    href = name_tag['href']
                    params = urllib.parse.parse_qs(urllib.parse.urlparse(href).query)
                    item_code = params.get('pno', [''])[0]
                    if item_code in seen_items:
                        continue
                    seen_items.add(item_code)

                    img_tag = product.select_one("div.img a img")
                    img_src = img_tag['src'] if img_tag else None
                    name = name_tag.get_text(strip=True)

                    price_old_tag = product.select_one("ul.prc .normal_prc")
                    price_discounted_tag = product.select_one("ul.prc strong")
                    price_old = price_old_tag.get_text(strip=True).replace(",", "").replace("원", "") if price_old_tag else None
                    price_discounted = price_discounted_tag.get_text(strip=True).replace(",", "").replace("원", "") if price_discounted_tag else None

                    if price_discounted:
                        price_discounted_int = int(price_discounted)
                        price = round(price_discounted_int * 1.15 / 1250, 2)
                        cena_na_site = round(price_discounted_int * 1.3 / 1250, 2)
                    else:
                        price = None
                        cena_na_site = None
                                             # ✅ Преобразуем в строку с десятичной точкой
                    cena_na_site_str = f"{cena_na_site:.2f}".replace(",", ".")
                    price_str = f"{price:.2f}".replace(",", ".")
                    name_en = translate_name_to_en(name)



                    moq = None
                    quantity_avail = None
                    category = assign_category(name)

                    ws.append([
                        img_src, brand_column, name,name_en, 'ea', moq, "20",
                        price_discounted, price_old, item_code,
                        f'=W{ws.max_row}&"/"&B{ws.max_row}', category, "ru",
                        20, "Все", 3, brand_name_en, price_str, cena_na_site_str, "Status", 'A'
                    ])
                except Exception as e:
                    print("Ошибка при разборе товара:", e)

    driver.quit()

    from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import tempfile

user_data_dir = tempfile.mkdtemp()
options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument(f"--user-data-dir={user_data_dir}")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

    # ---------------- Сохранение и загрузка на Google Drive -----------------
    file_path = f"beautydome.xlsx"
    wb.save(file_path)
    print(f"Файл локально сохранен: {file_path}")
    link = upload_to_drive(file_path, folder_id="10J-E4RcBJFfrdcqU_UAWask8BKTZ5Mw2")
    print(f"Ссылка на Google Drive: {link}")

# ---------------- Запуск -----------------
login_and_scrape("beelifecos","lapulik1983*")
