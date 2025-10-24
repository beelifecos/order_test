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
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import os
import pickle
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
    "BR000537": "AXIS-Y",
    "BR001045": "B_LAB",
    "BR000012": "Banila co",
    "BR000467": "Barr",
    "BR000373": "BARULAB",
    "BR000566": "BB LAB",
    "BR000549": "Be The Skin",
    "BR000498": "Beauty Recipe",
    "BR000486": "BEIGIC",
    "BR000013": "belif",
    "BR000395": "BellaMonster",
    "BR000188": "BENTON",
    "BR001010": "Bewants",
    "BR001149": "BIODANCE",
    "BR000837": "Biohealboh",
    "BR000664": "BONAJOUR",
    "BR000645": "BOTO",
    "BR000248": "Bouquetgarni",
    "BR001146": "BR MUD",
    "BR001221": "BRAYE",
    "BR000313": "briskin",
    "BR000506": "BUENO",
    "BR000629": "by : OUR",
    "BR000364": "BY ECOM",
    "BR000377": "CAILYN",
    "BR000445": "CANDYLAB",
    "BR000528": "celimax",
    "BR000435": "Cellfusion C",
    "BR000368": "Centellian24",
    "BR000808": "CHANGE FIT",
    "BR000555": "Chasin' Rabbits",
    "BR000531": "Chosungah Beauty",
    "BR000559": "CICATRI",
    "BR000881": "CIELO",
    "BR000190": "Ciracle",
    "BR001144": "CJ InnerB",
    "BR000049": "claires",
    "BR000084": "CLIO",
    "BR001314": "ClearDea",
    "BR000297": "CNP Cosmetics",
    "BR000066": "Coreana",
    "BR000236": "CORINGCO",
    "BR000447": "COSMETEA",
    "BR000607": "COSNORI",
    "BR001293": "colorgram",
    "BR001318": "Coralhaze",
    "BR000369": "CP-1",
    "BR000638": "DANONGWON",
    "BR000513": "Dasique",
    "BR000472": "Derma Maison",
    "BR000433": "DERMA:B",
    "BR000434": "DERMATORY",
    "BR001275": "DEARMAY",
    "BR000083": "Dewytree",
    "BR001277": "delphyr",
    "BR000994": "Dinto",
    "BR000339": "double dare",
    "BR000384": "DPC",
    "BR000149": "Dr.Althea",
    "BR000508": "Dr.ato",
    "BR000489": "Dr.Ceuracle",
    "BR000873": "Dr.FORHAIR",
    "BR000018": "Dr.Jart+",
    "BR000882": "Dr.PRIO",
    "BR000656": "Dr.WIN",
    "BR000432": "DUFT&DOFT",
    "BR000455": "E NATURE",
    "BR000381": "easybeauty",
    "BR000478": "EASYDEW",
    "BR000567": "ECOWINDY",
    "BR000580": "EDGE U",
    "BR000451": "EIIO",
    "BR000429": "ELENSILIA",
    "BR000041": "Elizavecca",
    "BR000352": "ELROEL",
    "BR000430": "ENOUGH",
    "BR000564": "espoir",
    "BR000807": "Essential",
    "BR001232": "ESTHER FORMULA",
    "BR000840": "EVER VITA",
    "BR001299": "EITHER AND",
    "BR000822": "EYECROWN",
    "BR000600": "EZWELL",
    "BR000505": "Farm stay",
    "BR000479": "Fiala Miji",
    "BR000492": "Fiera",
    "BR001030": "FOODOLOGY",
    "BR000606": "Formal Bee",
    "BR001096": "FRANKLY",
    "BR000520": "FREP",
    "BR000880": "FromBio",
    "BR000450": "Fromxoy",
    "BR001017": "Fullight",
    "BR000267": "G9",
    "BR000495": "GD11",
    "BR000518": "GILIM INTERNATIONAL",
    "BR001018": "Gleer",
    "BR000601": "Good Manner",
    "BR001003": "GOUTTER",
    "BR000416": "GROUNDPLAN",
    "BR001231": "GROWUS",
    "BR000577": "Gyeol Collagen",
    "BR000988": "hakit",
    "BR000026": "Hanyul",
    "BR000456": "Haruharu Wonder",
    "BR000439": "hddn lab",
    "BR000584": "Healing Bird",
    "BR000562": "Heart Percent",
    "BR000208": "Heimish",
    "BR000028": "HERA",
    "BR000504": "Herbnote",
    "BR001276": "hetras",
    "BR001234": "HEVEBLUE",
    "BR000022": "Holika Holika",
    "BR001075": "House of Hur",
    "BR000273": "Huxley",
    "BR001069": "HYAAH",
    "BR000474": "Hydrogen",
    "BR000310": "HYGGEE",
    "BR000540": "I DEW CARE",
    "BR001013": "IBL",
    "BR000077": "ILLIYOON",
    "BR000380": "I'm Sorry For My Skin",
    "BR000003": "Innisfree",
    "BR000005": "IOPE",
    "BR000458": "ISNTREE",
    "BR000535": "ISOI",
    "BR000023": "It's Skin",
    "BR000315": "IUNIK",
    "BR000672": "Jaunkyeol",
    "BR001071": "JAVIN DE SEOUL",
    "BR000140": "Jayjun",
    "BR000623": "JelliFit",
    "BR000573": "Jenny House",
    "BR000341": "JMsolution",
    "BR000529": "J'S DERMA",
    "BR000914": "Julie's Choice",
    "BR000296": "Jumiso",
    "BR001116": "JUST NATURE",
    "BR000582": "KAHI",
    "BR000897": "KAINE",
    "BR001104": "KAJA",
    "BR000634": "KANU",
    "BR000366": "KEEP COOL",
    "BR000459": "KIMJEONGMOON-ALOE",
    "BR000847": "Kirsh Blending",
    "BR000604": "KIUKIMIUM",
    "BR000321": "KLAVUU",
    "BR000218": "KOELF",
    "BR000599": "Korea Red Ginseng",
    "BR000585": "Kosette",
    "BR001191": "KSECRET",
    "BR001005": "Kwailnara",
    "BR000206": "Labiotte",
    "BR001014": "LACTONIA",
    "BR000358": "LAGOM",
    "BR000437": "lalaChuu",
    "BR000949": "LALARECIPE",
    "BR000004": "Laneige",
    "BR000019": "Leaders Insolution",
    "BR000186": "LEMONA",
    "BR000913": "Lifepharm",
    "BR000490": "Lilybyred",
    "BR000644": "LINGTEA",
    "BR000378": "LIZK",
    "BR000612": "LIZVIEW",
    "BR000603": "Lookas9",
    "BR000545": "Looks&Meii",
    "BR001101": "LYLA",
    "BR000469": "ma:nyo",
    "BR000511": "MADECA DERMA",
    "BR000651": "Mallingbooth",
    "BR000045": "Mamonde",
    "BR000667": "MARICEEL",
    "BR000574": "Mary&May",
    "BR000579": "MASIL",
    "BR000611": "Maxim",
    "BR000060": "Mediheal",
    "BR000362": "MediPeel",
    "BR000210": "MeFactory",
    "BR000978": "Melixir",
    "BR000153": "Memebox",
    "BR000371": "MERBLISS",
    "BR000388": "Merzy",
    "BR000507": "MIGUHARA",
    "BR000620": "Milk Touch",
    "BR000287": "MineralBio",
    "BR000616": "MINIMUM",
    "BR000657": "MIRACLE M",
    "BR000031": "MiseEnScene",
    "BR000015": "Missha",
    "BR000144": "Mizon",
    "BR000468": "MLB",
    "BR000284": "moonshot",
    "BR000617": "Mude",
    "BR000515": "MULAWEAR",
    "BR000660": "MY1CART",
    "BR001102": "myFORMULA",
    "BR000320": "NACIFIC",
    "BR000263": "NAKE UP FACE",
    "BR000813": "NAMING",
    "BR000548": "NARD",
    "BR000010": "Nature Republic",
    "BR001133": "Needly",
    "BR000205": "Neogen",
    "BR001145": "NewTree",
    "BR000576": "NINE LESS",
    "BR000647": "Nutri D-Day",
    "BR000815": "Ogi",
    "BR000304": "Olivarrier",
    "BR000538": "One-day's you",
    "BR000591": "Ongredients",
    "BR000839": "P.CALM",
    "BR000322": "Pack age",
    "BR000207": "Paparecipe",
    "BR000033": "Peripera",
    "BR000043": "Petitfee",
    "BR000618": "phykology",
    "BR000640": "PICKYWICKY",
    "BR000286": "plu",
    "BR000514": "PRAMY",
    "BR000476": "Preange",
    "BR000183": "Primera",
    "BR000597": "Pulmuone",
    "BR001004": "Puremay",
    "BR000247": "Pyunkang yul",
    "BR001100": "RaNiq",
    "BR000602": "RAWEL",
    "BR000232": "RE:P",
    "BR000385": "Real Barrier",
    "BR000329": "ROVECTIN",
    "BR000025": "Ryo",
    "BR000581": "SAFEAIR",
    "BR001011": "ScalpMed",
    "BR000475": "SCINIC",
    "BR000546": "Secret:X",
    "BR000178": "SecretKey",
    "BR000568": "SERUMKIND",
    "BR000869": "SERY BOX",
    "BR000671": "SHAKE BABY",
    "BR001103": "simplyO",
    "BR000078": "Skin1004",
    "BR000017": "Skinfood",
    "BR000503": "SKINRx LAB",
    "BR000048": "SNP",
    "BR000223": "So natural",
    "BR000330": "SOMEBYMI",
    "BR000195": "SON & PARK",
    "BR000443": "SOON MAMA",
    "BR000614": "SOON+",
    "BR000285": "SRB",
    "BR000396": "Style by Aiahn",
    "BR000007": "SU:M37˚",
    "BR000002": "Sulwhasoo",
    "BR001073": "Sungboon Editor",
    "BR000390": "Suntique",
    "BR000453": "SUR.MEDIC",
    "BR000569": "SUREBASE",
    "BR000212": "SWANICOCO",
    "BR000578": "TEAZEN",
    "BR000441": "TENZERO",
    "BR001222": "LAKA",
    "BR000437": "lalaChuu",
    "BR000949": "LALARECIPE",
    "BR000004": "Laneige",
    "BR000019": "Leaders Insolution",
    "BR000186": "LEMONA",
    "BR000913": "Lifepharm",
    "BR000490": "Lilybyred",
    "BR000644": "LINGTEA",
    "BR001183": "LINDSAY",
    "BR000378": "LIZK",
    "BR000612": "LIZVIEW",
    "BR000603": "Lookas9",
    "BR000545": "Looks&Meii",
    "BR001101": "LYLA",
    "BR000469": "ma:nyo",
    "BR000511": "MADECA DERMA",
    "BR000651": "Mallingbooth",
    "BR000045": "Mamonde",
    "BR000667": "MARICEEL",
    "BR000574": "Mary&May",
    "BR000579": "MASIL",
    "BR000611": "Maxim",
    "BR000060": "Mediheal",
    "BR000362": "MediPeel",
    "BR000210": "MeFactory",
    "BR000978": "Melixir",
    "BR000153": "Memebox",
    "BR001278": "MENOKIN",
    "BR000371": "MERBLISS",
    "BR000388": "Merzy",
    "BR000507": "MIGUHARA",
    "BR000620": "Milk Touch",
    "BR000287": "MineralBio",
    "BR000616": "MINIMUM",
    "BR000657": "MIRACLE M",
    "BR000031": "MiseEnScene",
    "BR000015": "Missha",
    "BR000144": "Mizon",
    "BR000468": "MLB",
    "BR000284": "moonshot",
    "BR000617": "Mude",
    "BR000515": "MULAWEAR",
    "BR000660": "MY1CART",
    "BR001102": "myFORMULA",
    "BR000320": "NACIFIC",
    "BR000263": "NAKE UP FACE",
    "BR000813": "NAMING",
    "BR000548": "NARD",
    "BR000010": "Nature Republic",
    "BR001290": "nesh",
    "BR001133": "Needly",
    "BR000205": "Neogen",
    "BR001145": "NewTree",
    "BR000576": "NINE LESS",
    "BR000647": "Nutri D-Day",
    "BR001230": "NONOER",
    "BR000262": "nooni",
    "BR001237": "ODDTYPE",
    "BR001199": "OOTD BEAUTY",
    "BR000815": "Ogi",
    "BR000304": "Olivarrier",
    "BR000538": "One-day's you",
    "BR000591": "Ongredients",
    "BR000839": "P.CALM",
    "BR000322": "Pack age",
    "BR000207": "Paparecipe",
    "BR000033": "Peripera",
    "BR000043": "Petitfee",
    "BR000618": "phykology",
    "BR000640": "PICKYWICKY",
    "BR000286": "plu",
    "BR000514": "PRAMY",
    "BR000476": "Preange",
    "BR000183": "Primera",
    "BR000597": "Pulmuone",
    "BR001004": "Puremay",
    "BR001223": "PURCELL",
    "BR000247": "Pyunkang yul",
    "BR001100": "RaNiq",
    "BR000602": "RAWEL",
    "BR000232": "RE:P",
    "BR000385": "Real Barrier",
    "BR001268": "RETURNITY",
    "BR000329": "ROVECTIN",
    "BR001287": "ROOTON",
    "BR000025": "Ryo",
    "BR000581": "SAFEAIR",
    "BR001011": "ScalpMed",
    "BR000475": "SCINIC",
    "BR000546": "Secret:X",
    "BR000178": "SecretKey",
    "BR000568": "SERUMKIND",
    "BR000869": "SERY BOX",
    "BR001282": "seapuri",
    "BR000671": "SHAKE BABY",
    "BR001235": "SHAISHAISHAI",
    "BR001103": "simplyO",
    "BR000078": "Skin1004",
    "BR000017": "Skinfood",
    "BR000503": "SKINRx LAB",
    "BR001242": "slowpure",
    "BR000048": "SNP",
    "BR000223": "So natural",
    "BR000330": "SOMEBYMI",
    "BR000195": "SON & PARK",
    "BR000443": "SOON MAMA",
    "BR000614": "SOON+",
    "BR000285": "SRB",
    "BR000396": "Style by Aiahn",
    "BR001194": "STUDIO 17",
    "BR000007": "SU:M37˚",
    "BR000002": "Sulwhasoo",
    "BR001073": "Sungboon Editor",
    "BR000390": "Suntique",
    "BR000453": "SUR.MEDIC",
    "BR000569": "SUREBASE",
    "BR000212": "SWANICOCO",
    "BR000578": "TEAZEN",
    "BR000441": "TENZERO",
    "BR000812": "TFIT",
    "BR000008": "The History of Whoo",
    "BR000654": "THE LAB by blanc doux",
    "BR000440": "THE MASK SHOP",
    "BR001130": "THE ORDINARY",
    "BR000454": "The Plant Base (P'lab)",
    "BR000006": "THEFACESHOP",
    "BR000280": "TIAM",
    "BR000883": "TIRTIR",
    "BR001301": "tiptoe",
    "BR001118": "Toi:L",
    "BR000011": "Tonymoly",
    "BR000016": "Too Cool For School",
    "BR000331": "TOSOWOONG",
    "BR000282": "touch in SOL",
    "BR000438": "Touch My body",
    "BR000519": "TOUN28",
    "BR000270": "TROIAREUKE",
    "BR001243": "Treecell",
    "BR000534": "Twinkle Pop",
    "BR000854": "UNOVE",
    "BR000431": "URANG",
    "BR000064": "VDL",
    "BR000392": "VELVIZO",
    "BR000470": "VIVLAS",
    "BR000307": "VT COSMETICS",
    "BR000200": "W.DRESSROOM",
    "BR000444": "WELLAGE",
    "BR000356": "WellDerma",
    "BR000289": "WHAMISA",
    "BR000258": "WonderBath",
    "BR000826": "Woori Nurungji",
    "BR000589": "Xpoiled",
    "BR000586": "YADAH",
    "BR000613": "YUNJAC",
    "BR000277": "ZYMOGEN",
    "BR001272": "z+piderm",
    "BR000609": "다신샵",
    "BR000615": "상아제약",
    "BR000643": "써니사이드수프 (SunnysideSoop)",
    "BR000610": "요뽀끼",
    "BR000631": "일광제과",
    "BR001134": "지알엔플러스",
    "BR000646": "컬러랩",
    "BR000666": "하움"
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
    options.add_argument('--headless')  # для GitHub Actions
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--user-data-dir=/tmp/selenium_unique')  # уникальный временный каталог

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    driver.get("https://wholesale.stylekorean.com/Member/SignIn")
    handle_alert(driver)

    driver.find_element(By.ID, "user_id").send_keys(username)
    driver.find_element(By.ID, "pwd").send_keys(password)
    driver.find_element(By.CSS_SELECTOR, ".Btn_Login[type='submit']").click()
    handle_alert(driver)

    # дальше код как у тебя…


    wb = Workbook()
    ws = wb.active
    ws.append([
        "Изображение", "Бренд", "Наименование", "Категория", "Единица измерения",
        "MOQ", "Фактический остаток", "in box", "Артикул", "Product code",
        "Цена Discounted KRW", "Cena na site $", "Price", "Language", "Lower limit",
        "User group", "Особенности", "Старая цена KRW","status","category","procent"
    ])

    file_path = '/Users/tyantamara/parser_stas_final_2.xlsx'

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
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000067", # ACWELL
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000473", # AESTURA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000457", # AHEADS
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000487", # AIRIVE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000811", # AKF
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000502", # ALETHEIA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001097", # ALLIONE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000081", # Amos
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000365", # AMPLE N
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000572", # AMTS (All My Things)
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000659", # AMUSE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000563", # And:ar
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000522", # ANN 365
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000516", # ANUA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000181", # Apieu
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001129", # APLB
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000152", # APRIL SKIN
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000294", # aromatica
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000625", # ATHINGS
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000367", # ATOPALM
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000558", # ATVT
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000301", # Avajar
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000537", # AXIS-Y
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001045", # B_LAB
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000012", # Banila co
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000467", # Barr
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000373", # BARULAB
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000566", # BB LAB
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000549", # Be The Skin
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000389", # Beauty of Joseon
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000498", # Beauty Recipe
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000486", # BEIGIC
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000013", # belif
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000395", # BellaMonster
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000188", # BENTON
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001010", # Bewants
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001149", # BIODANCE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000837", # Biohealboh
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000664", # BONAJOUR
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000645", # BOTO
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000248", # Bouquetgarni
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001146", # BR MUD
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001221", # BRAYE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000506", # BUENO
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000629", # by : OUR
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000364", # BY ECOM
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000377", # CAILYN
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000445", # CANDYLAB
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000528", # celimax
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001282", # seapuri
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000435", # Cellfusion C
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000368", # Centellian24
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000808", # CHANGE FIT
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000555", # Chasin' Rabbits
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000531", # Chosungah Beauty
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000559", # CICATRI
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000881", # CIELO
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000190", # Ciracle
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001144", # CJ InnerB
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000049", # claires
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000084", # CLIO
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000297", # CNP Cosmetics
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000066", # Coreana
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000236", # CORINGCO
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000447", # COSMETEA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000607", # COSNORI
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000189", # COSRX
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001293", # colorgram
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001318", # Coralhaze

"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000369", # CP-1
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000638", # DANONGWON
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000513", # Dasique
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000472", # Derma Maison
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000433", # DERMA:B
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000434", # DERMATORY
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001275", # DEARMAY
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000083", # Dewytree
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001277", # delphyr
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000994", # Dinto
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000221", # Doctor.G
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000339", # double dare
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000384", # DPC
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000149", # Dr.Althea
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000508", # Dr.ato
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000489", # Dr.Ceuracle
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000873", # Dr.FORHAIR
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000018", # Dr.Jart+
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000882", # Dr.PRIO
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000656", # Dr.WIN
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000432", # DUFT&DOFT
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000455", # E NATURE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001299", # EITHER AND
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001232", # ESTHER FORMULA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000381", # easybeauty
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000478", # EASYDEW
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000567", # ECOWINDY
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000580", # EDGE U
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000451", # EIIO
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000429", # ELENSILIA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000041", # Elizavecca
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000352", # ELROEL
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000430", # ENOUGH
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000564", # espoir
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000807", # Essential
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000840", # EVER VITA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000822", # EYECROWN
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000600", # EZWELL
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000505", # Farm stay
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000479", # Fiala Miji
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000492", # Fiera
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001030", # FOODOLOGY
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000606", # Formal Bee
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001096", # FRANKLY
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000520", # FREP
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000880", # FromBio
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000450", # Fromxoy
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001017", # Fullight
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000267", # G9
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000495", # GD11
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000518", # GILIM INTERNATIONAL
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001018", # Gleer
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000601", # Good Manner
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000168", # Goodal
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001003", # GOUTTER
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001231", # GROWUS
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000416", # GROUNDPLAN
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000577", # Gyeol Collagen
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000988", # hakit
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000026", # Hanyul
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001313", # hanskin
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000456", # Haruharu Wonder
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000439", # hddn lab
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000584", # Healing Bird
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000562", # Heart Percent
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000208", # Heimish
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000028", # HERA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000504", # Herbnote
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001234", # HEVEBLUE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000022", # Holika Holika
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001075", # House of Hur
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000273", # Huxley
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001069", # HYAAH
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000474", # Hydrogen
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000310", # HYGGEE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001276", # hetras
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000540", # I DEW CARE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001013", # IBL
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000077", # ILLIYOON
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000380", # I'm Sorry For My Skin
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000003", # Innisfree
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000005", # IOPE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000458", # ISNTREE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000535", # ISOI
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000023", # It's Skin
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000315", # IUNIK
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000672", # Jaunkyeol
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001071", # JAVIN DE SEOUL
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000140", # Jayjun
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000623", # JelliFit
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000573", # Jenny House
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000341", # JMsolution
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000529", # J'S DERMA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000914", # Julie's Choice
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000296", # Jumiso
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001116", # JUST NATURE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000582", # KAHI
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000897", # KAINE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001104", # KAJA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000634", # KANU
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000366", # KEEP COOL
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000459", # KIMJEONGMOON-ALOE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000847", # Kirsh Blending
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000604", # KIUKIMIUM
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000321", # KLAVUU
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001191", # KSECRET
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000218", # KOELF
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000599", # Korea Red Ginseng
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000585", # Kosette
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001005", # Kwailnara
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001222", # LAKA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000206", # Labiotte
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001014", # LACTONIA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000358", # LAGOM
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000437", # lalaChuu
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000949", # LALARECIPE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000004", # Laneige
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000019", # Leaders Insolution
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000186", # LEMONA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000913", # Lifepharm
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000490", # Lilybyred
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001183", # LINDSAY
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000644", # LINGTEA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000378", # LIZK
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000612", # LIZVIEW
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000603", # Lookas9
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000545", # Looks&Meii
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001101", # LYLA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000469", # ma:nyo
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000511", # MADECA DERMA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000651", # Mallingbooth
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000045", # Mamonde
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000667", # MARICEEL
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000574", # Mary&May
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000579", # MASIL
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000611", # Maxim
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000060", # Mediheal
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000362", # MediPeel
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000210", # MeFactory
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000978", # Melixir
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000153", # Memebox
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001278", # MENOKIN
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000371", # MERBLISS
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000388", # Merzy
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000507", # MIGUHARA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000620", # Milk Touch
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000287", # MineralBio
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000616", # MINIMUM
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000657", # MIRACLE M
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000031", # MiseEnScene
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000015", # Missha
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000872", # MIXSOON
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000144", # Mizon
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000468", # MLB
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000284", # moonshot
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000617", # Mude
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000515", # MULAWEAR
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000660", # MY1CART
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001102", # myFORMULA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000320", # NACIFIC
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000263", # NAKE UP FACE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000813", # NAMING
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000548", # NARD
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000010", # Nature Republic
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001290", # nesh
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001133", # Needly
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000205", # Neogen
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001145", # NewTree
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000576", # NINE LESS
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000647", # Nutri D-Day
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001230", # NONOER
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000262", # nooni
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001237", # ODDTYPE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001199", # OOTD BEAUTY
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000815", # Ogi
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000304", # Olivarrier
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000538", # One-day's you
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000591", # Ongredients
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000839", # P.CALM
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000322", # Pack age
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000207", # Paparecipe
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000033", # Peripera
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000043", # Petitfee
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000618", # phykology
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000640", # PICKYWICKY
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000286", # plu
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000514", # PRAMY
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000476", # Preange
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000183", # Primera
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000597", # Pulmuone
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001004", # Puremay
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001223", # PURCELL
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000247", # Pyunkang yul
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001100", # RaNiq
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000602", # RAWEL
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000232", # RE:P
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001268", # RETURNITY
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000385", # Real Barrier
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000317", # rom&nd
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001287", # ROOTON
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000329", # ROVECTIN
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000025", # Ryo
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000581", # SAFEAIR
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001011", # ScalpMed
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000475", # SCINIC
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000546", # Secret:X
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000178", # SecretKey
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000568", # SERUMKIND
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000869", # SERY BOX
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001235", # SHAISHAISHAI
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000671", # SHAKE BABY
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001103", # simplyO
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000078", # Skin1004
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000017", # Skinfood
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000503", # SKINRx LAB
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001242", # slowpure
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000048", # SNP
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000223", # So natural
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000330", # SOMEBYMI
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000195", # SON & PARK
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000443", # SOON MAMA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000614", # SOON+
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000285", # SRB
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000396", # Style by Aiahn
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001194", # STUDIO 17
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000007", # SU:M37˚
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000002", # Sulwhasoo
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001073", # Sungboon Editor
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000390", # Suntique
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000453", # SUR.MEDIC
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000569", # SUREBASE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000212", # SWANICOCO
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000578", # TEAZEN
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000441", # TENZERO
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000812", # TFIT
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000008", # The History of Whoo
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000654", # THE LAB by blanc doux
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000440", # THE MASK SHOP
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001130", # THE ORDINARY
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000454", # The Plant Base (P'lab)
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000006", # THEFACESHOP
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000280", # TIAM
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000883", # TIRTIR
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001301", # tiptoe
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000627", # TOCOBO
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001118", # Toi:L
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000011", # Tonymoly
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000016", # Too Cool For School
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000331", # TOSOWOONG
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000282", # touch in SOL
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000438", # Touch My body
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000519", # TOUN28
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000270", # TROIAREUKE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001243", # Treecell
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000534", # Twinkle Pop
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000854", # UNOVE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000431", # URANG
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000064", # VDL
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000392", # VELVIZO
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000470", # VIVLAS
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000307", # VT COSMETICS
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000200", # W.DRESSROOM
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000444", # WELLAGE
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000356", # WellDerma
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000289", # WHAMISA
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000258", # WonderBath
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000826", # Woori Nurungji
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000589", # Xpoiled
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000586", # YADAH
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000613", # YUNJAC
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000277", # ZYMOGEN
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001272", # z+piderm
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000609", # 다신샵
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000615", # 상아제약
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000643", # 써니사이드수프 (SunnysideSoop)
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000610", # 요뽀끼
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000631", # 일광제과
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR001134", # 지알엔플러스
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000646", # 컬러랩
"https://wholesale.stylekorean.com/Product/BrandProduct?brand_cd=BR000666", # 하움
    ]

    for brand_url in brand_urls:
        brand_name = extract_brand_name(brand_url)
        print(f"Scraping products for brand: {brand_name}")

        driver.get(brand_url)
        handle_alert(driver)
        WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.CLASS_NAME, "album")))

        page_links = driver.find_elements(By.CLASS_NAME, "page-link")
        num_pages = len(page_links) if page_links else 1
        if len(page_links) >= 3:
            num_pages_element = page_links[-3]
            num_pages_label = num_pages_element.get_attribute("aria-label")
            if num_pages_label:
                num_pages = int(num_pages_label.split()[-1])

        for page_num in range(1, num_pages + 1):
            handle_alert(driver)
            WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.CLASS_NAME, "album")))
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
                     # ✅ Преобразуем в строку с десятичной точкой
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

            # --- Загружаем (обновляем) на Google Drive ---
            try:
                # Ищем существующий файл по имени
                query = f"'{folder_id}' in parents and trashed=false and title='{file_name}'"
                file_list = drive.ListFile({'q': query}).GetList()

                if file_list:
                    # ✅ Файл найден — обновляем
                    file_drive = file_list[0]
                    print(f"Найден существующий файл '{file_name}', обновляем содержимое...")
                else:
                    # ❗ Файл не найден — создаём новый
                    file_drive = drive.CreateFile({'title': file_name, 'parents': [{'id': folder_id}]})
                    print(f"Файл '{file_name}' не найден, создаём новый...")

                # Загружаем файл
                file_drive.SetContentFile(file_path)
                file_drive.Upload()
                print(f"✅ Файл '{file_name}' успешно обновлён на Google Drive после страницы {page_num}")

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
login_and_scrape("beelifecos", "1983beelif")
