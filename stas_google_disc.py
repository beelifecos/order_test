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
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import os

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
    "BR000642": "12GRABS",
    "BR000257": "16Brand",
    "BR000243": "23 years old",
    "BR000014": "3 Concept Eyes [3CE]",
    "BR000357": "9Wishes",
    "BR000609": "다신샵",
    "BR001377": "듀얼소닉",
    "BR000615": "상아제약",
    "BR000643": "써니사이드수프 (SunnysideSoop)",
    "BR000610": "요뽀끼",
    "BR000631": "일광제과",
    "BR000646": "컬러랩",
    "BR000666": "하움",
    "BR000142": "A.H.C",
    "BR000590": "A+ Clean Up",
    "BR001115": "ABEREDE",
    "BR000311": "Abib",
    "BR000080": "About me",
    "BR001368": "ABOUT TONE",
    "BR000482": "AcroPass",
    "BR000067": "ACWELL",
    "BR000473": "AESTURA",
    "BR000457": "AHEADS",
    "BR000487": "AIRIVE",
    "BR000811": "AKF",
    "BR000502": "ALETHEIA",
    "BR001097": "ALLIONE",
    "BR001148": "ALTERNATIVESTEREO",
    "BR001295": "Amazin' Graze",
    "BR000081": "Amos",
    "BR000365": "AMPLE N",
    "BR000572": "AMTS (All My Things)",
    "BR000659": "AMUSE",
    "BR000563": "And:ar",
    "BR000669": "Ando",
    "BR000522": "ANN 365",
    "BR001239": "Another Face",
    "BR000516": "ANUA",
    "BR001302": "AOU",
    "BR000181": "Apieu",
    "BR001129": "APLB",
    "BR000152": "APRIL SKIN",
    "BR001206": "ARENCIA",
    "BR001037": "Ariul",
    "BR000294": "aromatica",
    "BR001364": "arwe",
    "BR001213": "ASIS-TOBE",
    "BR000625": "ATHINGS",
    "BR000367": "ATOPALM",
    "BR000558": "ATVT",
    "BR000301": "Avajar",
    "BR000537": "AXIS-Y",
    "BR001045": "B:LAB",
    "BR000012": "Banila co",
    "BR001138": "BAREN",
    "BR000467": "Barr",
    "BR000373": "BARULAB",
    "BR000566": "BB LAB",
    "BR001367": "BBIA",
    "BR000549": "Be The Skin",
    "BR001273": "BEAUND",
    "BR000389": "Beauty of Joseon",
    "BR000498": "Beauty Recipe",
    "BR000486": "BEIGIC",
    "BR000013": "belif",
    "BR000395": "BellaMonster",
    "BR000188": "BENTON",
    "BR001010": "Bewants",
    "BR001149": "BIODANCE",
    "BR001220": "Blessed Moon",
    "BR000199": "BLITHE",
    "BR001351": "Blood",
    "BR001177": "BOM",
    "BR000664": "BONAJOUR",
    "BR000645": "BOTO",
    "BR000248": "Bouquetgarni",
    "BR001146": "BR MUD",
    "BR001221": "BRAYE",
    "BR000313": "briskin",
    "BR000506": "BUENO",
    "BR000629": "by : OUR",
    "BR000364": "BY ECOM",
    "BR001361": "by juccy",
    "BR000377": "CAILYN",
    "BR000445": "CANDYLAB",
    "BR001217": "CATCH ME PATCH",
    "BR001190": "CCAM BBAK",
    "BR001257": "CELDERMA",
    "BR000528": "celimax",
    "BR000435": "Cellfusion C",
    "BR000368": "Centellian24",
    "BR000808": "CHANGE FIT",
    "BR000555": "Chasin' Rabbits",
    "BR001219": "CHICSKIN",
    "BR000531": "Chosungah Beauty",
    "BR001354": "Chwi",
    "BR000559": "CICATRI",
    "BR000881": "CIELO",
    "BR000190": "Ciracle",
    "BR001144": "CJ InnerB",
    "BR001209": "CKD",
    "BR000049": "claires",
    "BR001314": "ClearDea",
    "BR000084": "CLIO",
    "BR001240": "CNP BYE ODTD",
    "BR000297": "CNP Cosmetics",
    "BR001293": "colorgram",
    "BR001318": "Coralhaze",
    "BR000066": "Coreana",
    "BR000236": "CORINGCO",
    "BR001081": "COS DE BAHA",
    "BR000447": "COSMETEA",
    "BR000607": "COSNORI",
    "BR000189": "COSRX",
    "BR000369": "CP-1",
    "BR000334": "d'Alba",
    "BR001261": "Danahan",
    "BR000638": "DANONGWON",
    "BR000513": "Dasique",
    "BR001375": "Dear Doer",
    "BR001275": "DEARMAY",
    "BR001176": "DearMYDEW",
    "BR001277": "delphyr",
    "BR001255": "Derma block",
    "BR000472": "Derma Maison",
    "BR000433": "DERMA:B",
    "BR000434": "DERMATORY",
    "BR000083": "Dewytree",
    "BR000994": "Dinto",
    "BR001303": "DIXIONIST",
    "BR000221": "Doctor.G",
    "BR000339": "double dare",
    "BR000384": "DPC",
    "BR001323": "Dr. Reju-All",
    "BR000149": "Dr.Althea",
    "BR000508": "Dr.ato",
    "BR001346": "Dr.Bio",
    "BR000489": "Dr.Ceuracle",
    "BR001312": "Dr.CPU",
    "BR000873": "Dr.FORHAIR",
    "BR001284": "Dr.Groot",
    "BR000018": "Dr.Jart+",
    "BR001184": "Dr.Melaxin",
    "BR001307": "Dr.nineteen",
    "BR000882": "Dr.PRIO",
    "BR000656": "Dr.WIN",
    "BR001319": "Dropbe",
    "BR000432": "DUFT&amp;DOFT",
    "BR000455": "E NATURE",
    "BR000381": "easybeauty",
    "BR000478": "EASYDEW",
    "BR000567": "ECOWINDY",
    "BR000580": "EDGE U",
    "BR001370": "EDIT.B",
    "BR000451": "EIIO",
    "BR001299": "EITHER AND",
    "BR000429": "ELENSILIA",
    "BR000041": "Elizavecca",
    "BR000352": "ELROEL",
    "BR000430": "ENOUGH",
    "BR000564": "espoir",
    "BR001350": "essel",
    "BR000807": "Essential",
    "BR001232": "ESTHER FORMULA",
    "BR000001": "Etude",
    "BR000840": "EVER VITA",
    "BR001353": "Eyecandy",
    "BR000822": "EYECROWN",
    "BR000600": "EZWELL",
    "BR000505": "Farm stay",
    "BR001270": "FATION",
    "BR000479": "Fiala Miji",
    "BR000492": "Fiera",
    "BR001274": "FILFLO",
    "BR001030": "FOODOLOGY",
    "BR001349": "FORBEAUT",
    "BR000606": "Formal Bee",
    "BR001096": "FRANKLY",
    "BR000520": "FREP",
    "BR000880": "FromBio",
    "BR000450": "Fromxoy",
    "BR000481": "Frudia",
    "BR001017": "Fullight",
    "BR001227": "FULLY",
    "BR001161": "FWEE"    
    }
    return brand_name_map.get(brand_cd, "Unknown Brand")

def handle_alert(driver):
    try:
        WebDriverWait(driver, 3).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
    except:
        pass

# --- Основная функция скрапинга ---
def login_and_scrape(username, password):
    options = Options()
    options.add_argument('--disable-notifications')
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    driver.get("https://stylekoreankbeautywholesale.com/Member/SignIn")
    handle_alert(driver)

    driver.find_element(By.ID, "user_id").send_keys(username)
    driver.find_element(By.ID, "pwd").send_keys(password)
    driver.find_element(By.CSS_SELECTOR, ".Btn_Login[type='submit']").click()
    handle_alert(driver)

    wb = Workbook()
    ws = wb.active
    ws.append([
        "Изображение", "Бренд", "Наименование", "Категория", "Единица измерения",
        "MOQ", "Фактический остаток", "in box", "Артикул", "Product code",
        "Цена Discounted KRW", "Cena na site", "Price_opt", "Language", "Lower limit",
        "User group", "Особенности", "Старая цена KRW","status","category","procent", "Cena na site $","Price"
    ])

    file_path = '/Users/tyantamara/parser_stas_final_1.xlsx'
    file_name = "parser_stas_final_1.xlsx"


    # --- Google Drive ---
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)
    folder_id = "10J-E4RcBJFfrdcqU_UAWask8BKTZ5Mw2"

    # --- Список URL брендов ---
    brand_urls = [
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000642", #12GRABS
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000257", #16Brand
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000243", #23 years old
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000014", #3 Concept Eyes [3CE]
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000357", #9Wishes
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000609", #다신샵
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001377", #듀얼소닉
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000615", #상아제약
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000643", #써니사이드수프 (SunnysideSoop)
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000610", #요뽀끼
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000631", #일광제과
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000646", #컬러랩
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000666", #하움
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000142", #A.H.C
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000590", #A+ Clean Up
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001115", #ABEREDE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000311", #Abib
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000080", #About me
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001368", #ABOUT TONE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000482", #AcroPass
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000067", #ACWELL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000473", #AESTURA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000457", #AHEADS
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000487", #AIRIVE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000811", #AKF
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000502", #ALETHEIA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001097", #ALLIONE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001148", #ALTERNATIVESTEREO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001295", #Amazin' Graze
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000081", #Amos
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000365", #AMPLE N
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000572", #AMTS (All My Things)
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000659", #AMUSE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000563", #And:ar
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000669", #Ando
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000522", #ANN 365
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001239", #Another Face
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000516", #ANUA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001302", #AOU
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000181", #Apieu
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001129", #APLB
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000152", #APRIL SKIN
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001206", #ARENCIA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001037", #Ariul
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000294", #aromatica
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001364", #arwe
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001213", #ASIS-TOBE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000625", #ATHINGS
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000367", #ATOPALM
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000558", #ATVT
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000301", #Avajar
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000537", #AXIS-Y
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001045", #B:LAB
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000012", #Banila co
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001138", #BAREN
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000467", #Barr
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000373", #BARULAB
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000566", #BB LAB
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001367", #BBIA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000549", #Be The Skin
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001273", #BEAUND
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000389", #Beauty of Joseon
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000498", #Beauty Recipe
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000486", #BEIGIC
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000013", #belif
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000395", #BellaMonster
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000188", #BENTON
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001010", #Bewants
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001149", #BIODANCE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001220", #Blessed Moon
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000199", #BLITHE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001351", #Blood
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001177", #BOM
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000664", #BONAJOUR
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000645", #BOTO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000248", #Bouquetgarni
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001146", #BR MUD
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001221", #BRAYE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000313", #briskin
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000506", #BUENO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000629", #by : OUR
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000364", #BY ECOM
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001361", #by juccy
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000377", #CAILYN
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000445", #CANDYLAB
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001217", #CATCH ME PATCH
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001190", #CCAM BBAK
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001257", #CELDERMA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000528", #celimax
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000435", #Cellfusion C
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000368", #Centellian24
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000808", #CHANGE FIT
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000555", #Chasin' Rabbits
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001219", #CHICSKIN
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000531", #Chosungah Beauty
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001354", #Chwi
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000559", #CICATRI
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000881", #CIELO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000190", #Ciracle
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001144", #CJ InnerB
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001209", #CKD
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000049", #claires
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001314", #ClearDea
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000084", #CLIO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001240", #CNP BYE ODTD
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000297", #CNP Cosmetics
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001293", #colorgram
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001318", #Coralhaze
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000066", #Coreana
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000236", #CORINGCO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001081", #COS DE BAHA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000447", #COSMETEA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000607", #COSNORI
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000189", #COSRX
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000369", #CP-1
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000334", #d'Alba
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001261", #Danahan
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000638", #DANONGWON
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000513", #Dasique
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001375", #Dear Doer
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001275", #DEARMAY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001176", #DearMYDEW
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001277", #delphyr
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001255", #Derma block
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000472", #Derma Maison
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000433", #DERMA:B
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000434", #DERMATORY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000083", #Dewytree
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000994", #Dinto
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001303", #DIXIONIST
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000221", #Doctor.G
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000384", #double dare
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000384", #DPC
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001323", #Dr. Reju-All
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000149", #Dr.Althea
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000508", #Dr.ato
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001346", #Dr.Bio
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000489", #Dr.Ceuracle
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001312", #Dr.CPU
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000873", #Dr.FORHAIR
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001284", #Dr.Groot
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000018", #Dr.Jart+
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001184", #Dr.Melaxin
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001307", #Dr.nineteen
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000882", #Dr.PRIO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000656", #Dr.WIN
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001319", #Dropbe
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000432", #DUFT&DOFT
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000455", #E NATURE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000381", #easybeauty
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000478", #EASYDEW
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000567", #ECOWINDY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000580", #EDGE U
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000357", #EDIT.B
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001370", #EIIO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001299", #EITHER AND
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000429", #ELENSILIA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000041", #Elizavecca
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000352", #ELROEL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000430", #ENOUGH
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000564", #espoir
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001350", #essel
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000807", #Essential
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001232", #ESTHER FORMULA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000001", #Etude
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000840", #EVER VITA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001353", #Eyecandy
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000822", #EYECROWN
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000600", #EZWELL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000505", #Farm stay
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001270", #FATION
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000479", #Fiala Miji
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000492", #Fiera
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001274", #FILFLO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001030", #FOODOLOGY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001349", #FORBEAUT
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000606", #Formal Bee
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001096", #FRANKLY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000520", #FREP
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000880", #FromBio
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000450", #Fromxoy
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000481", #Frudia
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001017", #Fullight
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001227", #FULLY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001161", #FWEE
    ]

    for brand_url in brand_urls:
        brand_name = extract_brand_name(brand_url)
        print(f"Scraping products for brand: {brand_name}")

        driver.get(brand_url)
        handle_alert(driver)
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CLASS_NAME, "album")))

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
            handle_alert(driver)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "album")))
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            # --- Обновлённый и устойчивый поиск карточек ---
            product_cards = soup.find_all("div", class_="card mb-4 shadow-sm")
            for card in product_cards:
                try:
                    # --- IMG ---
                    img_element = card.find("img", class_="Img_Product")
                    img_src = img_element.get('src') if img_element else None

                    # --- BRAND ---
                    brand_el = card.find("span", class_="companyTxt")
                    brand_name_text = brand_el.get_text(strip=True) if brand_el else ""
                    brand_name_clean = brand_name_text.replace("[", "").replace("]", "")

                    # --- NAME ---
                    name_el = card.find("span", class_=["productTxt"])
                    name = name_el.get_text(strip=True) if name_el else ""
                    # --- FULL NAME ---
                    full_name = f"{brand_name_clean} {name}"



                    # --- CATEGORY (твоя функция) ---
                    category = assign_category(name)

                    # --- SKU / Артикул ---
                    item_code = ""
                    sku_el = card.find("span", class_="productCodeTxt")
                    if sku_el:
                        sku_text = sku_el.get_text(separator=" ").strip()
                        # пытаемся найти "SKU: <code>"
                        m = re.search(r"SKU:\s*([A-Za-z0-9\-]+)", sku_text)
                        if m:
                            item_code = m.group(1)
                        else:
                            # fallback: взять первое вхождение похожее на артикул
                            parts = sku_text.split()
                            for p in parts:
                                if re.match(r"[A-Za-z0-9\-]{4,}", p):
                                    item_code = p
                                    break

                    # --- BARCODE / Product code ---
                    Product_code = ""
                    bar_el = card.find("span", class_="barcodeTxt")
                    if bar_el:
                        bar_text = bar_el.get_text(strip=True)
                        m2 = re.search(r"[:]\s*([0-9A-Za-z\-]+)$", bar_text)
                        if m2:
                            Product_code = m2.group(1)
                        else:
                            # fallback remove label
                            Product_code = bar_text.replace("Bar Code:", "").strip()

                    # --- MOQ ---
                    moq = ""
                    moq_el = card.find("span", class_="moqTxt")
                    if moq_el:
                        moq_text = moq_el.get_text(strip=True)
                        m3 = re.search(r"(\d+)", moq_text)
                        if m3:
                            moq = m3.group(1)
                        else:
                            moq = moq_text.replace("MOQ:", "").replace("ea", "").strip()

                    # --- STOCK ---
                    quantity_availabl = ""
                    qty_el = card.find("span", class_="qtyTxt")
                    if qty_el:
                        qty_text = qty_el.get_text(strip=True)
                        qty_clean = qty_text.replace("ea", "").replace(",", "").strip()
                        # взять первое найденное число
                        m4 = re.search(r"(\d+)", qty_clean)
                        if m4:
                            quantity_availabl = m4.group(1)
                        else:
                            quantity_availabl = qty_clean

                    # --- IN BOX (pieces per box) ---
                    pieces_per_box = "20"
                    box_el = card.find("span", class_="boxCnt")
                    if box_el:
                        box_text = box_el.get_text(" ", strip=True)
                        m5 = re.search(r"(\d+)", box_text)
                        if m5:
                            pieces_per_box = m5.group(1)
                        else:
                            # если нет цифр, оставить дефолт
                            pieces_per_box = "20"

                    # --- DISCOUNTED PRICE ---
                    price_discounted = 0.0
                    # priceTxt встречается в нескольких местах; берём первый, который выглядит как цена с KRW
                    price_el = None
                    for p in card.find_all("span", class_="priceTxt"):
                        t = p.get_text(strip=True)
                        if "KRW" in t or re.search(r"\d", t):
                            price_el = p
                            break
                    if price_el:
                        price_text = price_el.get_text(strip=True)
                        price_clean = price_text.replace("KRW", "").replace(",", "").replace(".00", "").strip()
                        # взять число
                        m6 = re.search(r"(\d+)", price_clean)
                        if m6:
                            try:
                                price_discounted = float(m6.group(1))
                            except:
                                price_discounted = 0.0

                    # --- OLD PRICE ---
                    price_old_1 = f"=ROUNDUP(L{ws.max_row+1}*1500; -3)"  # для русской локали
                    price_old_el = card.find("span", class_="priceOld2")
                    if price_old_el:
                        old_text = price_old_el.get_text(strip=True)
                        old_clean = old_text.replace("KRW", "").replace(",", "").replace(".00", "").strip()
                        m7 = re.search(r"(\d+)", old_clean)
                        if m7:
                            try:
                                price_old = float(m7.group(1))
                            except:
                                price_old = None

                    # --- PRICE CALC ---
                    cena_na_site = round(price_discounted * 1.2 / 1250, 2) if price_discounted else 0
                    price = round(price_discounted * 1.1 / 1250, 2) if price_discounted else 0
                    procent = round(price_discounted / price_old, 2) if (price_old and price_old != 0) else 0
                    # --- Приводим цены к строкам с точкой ---
                    cena_na_site = str(cena_na_site).replace(",", ".")
                    price = str(price).replace(",", ".")

                    # --- STATUS ---
                    brand_for_status = brand_name_clean if brand_name_clean else brand_name
                    status_value = f"Бренд///{brand_for_status[:1].upper() if brand_for_status else 'X'}///{brand_for_status}"
                    STATUS = "A"

                    # --- CLEAN CODES ---
                    item_code_clean = re.sub(r'\s+', '', item_code) if item_code else ""
                    product_code_clean = re.sub(r'\s+', '', Product_code) if Product_code else ""

                    import math

                    # Словарь с коэффициентами для брендов
                    brand_discounts = {
                     "ALLIONE": 0.5,     # например, 50% от старой цены
                     "B:LAB": 0.5,   # например, 55% от старой цены
                     "Be The Skin": 0.15   # например, 55% от старой цены

                     }

                     # Проверяем, есть ли бренд в словаре и есть ли старая цена
                    if procent == 1 and brand_name_clean in brand_discounts and price_old:
                     discount = brand_discounts[brand_name_clean]
                     cena_na_site_1 = round(price_old * discount / 1250, 2)
                     price_1 = round(price_old * discount / 1250, 2)
                    else:
                       cena_na_site_1 = cena_na_site
                       price_1 = price


                    # --- APPEND ROW ---
                    ws.append([
                        img_src,
                        brand_name_clean if brand_name_clean else brand_name,
                        full_name,
                        category,
                        'ea',
                        moq,
                        quantity_availabl,
                        pieces_per_box,
                        item_code_clean,
                        product_code_clean,
                        price_discounted,
                        cena_na_site,
                        price,
                        'ru',
                        pieces_per_box,
                        'Все',
                        '1',
                        price_old_1,
                        STATUS,
                        status_value,
                        procent,
                        cena_na_site_1,
                        price_1

                    ])

                except Exception as e:
                    print("Error parsing product:", e)
                

            # --- Сохраняем Excel локально ---
            try:
                wb.save(file_path)
                print(f"File saved successfully after page {page_num}")
            except Exception as e:
                print("Error saving file:", e)

                     # --- Загружаем на Google Drive ---
            try:
                # Ищем файл в папке
                query = f"'{folder_id}' in parents and trashed=false and title='{file_name}'"
                file_list = drive.ListFile({'q': query}).GetList()

                if file_list:
                    # Файл существует — обновляем содержимое (ID сохраняется!)
                    file_drive = file_list[0]
                    print(f"Обновление файла на Google Drive: {file_drive['id']}")
                else:
                    # Файла нет — создаём новый
                    file_drive = drive.CreateFile({'title': file_name, 'parents':[{'id': folder_id}]})
                    print("Создаю новый файл на Google Drive")

                file_drive.SetContentFile(file_path)
                file_drive.Upload()

                print(f"Файл '{file_name}' успешно обновлён на Google Drive без смены ID")

            except Exception as e:
                print("Error uploading file to Google Drive:", e)

            # --- Переход на следующую страницу ---
            if page_num < num_pages:
                try:
                    next_page_button = driver.find_element(By.XPATH, f"//a[@class='page-link' and @page='{page_num + 1}']")
                    next_page_button.click()
                except Exception as e:
                    print("Error clicking next page:", e)
                    break

    driver.quit()
    print("Scraping completed.")

# --- Запуск парсера ---
login_and_scrape("beelifecos","1983beelif")
