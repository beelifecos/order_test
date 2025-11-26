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
    if any(k in name_lower for k in ["샴푸", "shampoo","왁싱 매니큐어","미쟝센","헤어커버","LPP 트리트","아르드포 스프레이","염색", "컨디셔너","일진 케론 시스테인 웨이브","퍼퓸 린스", "conditioner","아이 팔레트", "린스","트리트먼트","hair treatment", "헤어 린스","쿨링 토닉","케라틴", "헤어칼라","크리닉 칼라"," 헤어 칼라"," 헤어","스타일инг 무스 ","셋팅 스프레이", "hair", "treatment", "헤어팩","시스테인","헤어비비", "hair pack","새치", "헤어오일", "hair oil"]):
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
    "BR001161": "FWEE",
    "BR000267": "G9",
    "BR000495": "GD11",
    "BR000518": "GILIM INTERNATIONAL",
    "BR001185": "Ginger6",
    "BR001018": "Gleer",
    "BR001327": "glow",
    "BR000601": "Good Manner",
    "BR000168": "Goodal",
    "BR001229": "GOONGBE",
    "BR001003": "GOUTTER",
    "BR001345": "GRABITY",
    "BR001362": "GRAFEN",
    "BR000605": "Green Monster",
    "BR001134": "GRN Plus",
    "BR000416": "GROUNDPLAN",
    "BR001231": "GROWUS",
    "BR000577": "Gyeol Collagen",
    "BR001179": "HairPlus",
    "BR000988": "hakit",
    "BR001313": "hanskin",
    "BR000026": "Hanyul",
    "BR000456": "Haruharu Wonder",
    "BR001379": "HAUA",
    "BR000439": "hddn lab",
    "BR001373": "HealGrids",
    "BR000584": "Healing Bird",
    "BR000562": "Heart Percent",
    "BR000208": "Heimish",
    "BR000028": "HERA",
    "BR000504": "Herbnote",
    "BR001276": "hetras",
    "BR001234": "HEVEBLUE",
    "BR001165": "hince",
    "BR000022": "Holika Holika",
    "BR001075": "House of Hur",
    "BR001245": "HUECALM",
    "BR000273": "Huxley",
    "BR001069": "HYAAH",
    "BR000474": "Hydrogen",
    "BR000310": "HYGGEE",
    "BR000540": "I DEW CARE",
    "BR001013": "IBL",
    "BR001196": "ID PLACOSMETICS",
    "BR000077": "ILLIYOON",
    "BR001164": "Ilso",
    "BR001288": "I'm meme",
    "BR000380": "I'm Sorry For My Skin",
    "BR000003": "Innisfree",
    "BR000005": "IOPE",
    "BR000458": "ISNTREE",
    "BR000535": "ISOI",
    "BR000023": "I's Skin",
    "BR000315": "IUNIK",
    "BR001169": "jaminkyung",
    "BR000672": "Jaunkyeol",
    "BR001071": "JAVIN DE SEOUL",
    "BR000140": "Jayjun",
    "BR000623": "JelliFit",
    "BR000573": "Jenny House",
    "BR000341": "JMsolution",
    "BR000529": "J'S DERMA",
    "BR000914": "Julie's Choice",
    "BR001330": "JULYME",
    "BR000296": "Jumiso",
    "BR000665": "Jungsaemmool",
    "BR001306": "JUST I AM",
    "BR001116": "JUST NATURE",
    "BR000582": "KAHI",
    "BR000897": "KAINE",
    "BR001104": "KAJA",
    "BR000634": "KANU",
    "BR000366": "KEEP COOL",
    "BR001207": "KEYTH",
    "BR001335": "KILLIT",
    "BR000459": "KIMJEONGMOON-ALOE",
    "BR001324": "KINOHIMITSU",
    "BR000847": "Kirsh Blending",
    "BR001291": "Kitsui",
    "BR000604": "KIUKIMIUM",
    "BR000321": "KLAVUU",
    "BR001226": "KllureMY",
    "BR000218": "KOELF",
    "BR001236": "KOPHER",
    "BR000599": "Korea Red Ginseng",
    "BR000585": "Kosette",
    "BR001191": "KSECRET",
    "BR001317": "KUNDAL",
    "BR001005": "Kwailnara",
    "BR000206": "Labiotte",
    "BR001014": "LACTONIA",
    "BR000358": "LAGOM",
    "BR001222": "LAKA",
    "BR000437": "lalaChuu",
    "BR000949": "LALARECIPE",
    "BR001175": "lalucell",
    "BR000004": "Laneige",
    "BR000019": "Leaders Insolution",
    "BR000186": "LEMONA",
    "BR001357": "Libresse",
    "BR000913": "Lifepharm",
    "BR000490": "Lilybyred",
    "BR001343": "Lilyeve",
    "BR001183": "LINDSAY",
    "BR000644": "LINGTEA",
    "BR001322": "LINGTEA",
    "BR000378": "LIZK",
    "BR000612": "LIZVIEW",
    "BR001178": "lleafill",
    "BR000603": "Lookas9",
    "BR000545": "Looks&Meii",
    "BR001101": "LYLA",
    "BR000469": "ma:nyo",
    "BR001256": "Madeca 21",
    "BR000511": "MADECA DERMA",
    "BR000989": "Makep:rem",
    "BR000651": "Mallingbooth",
    "BR000045": "Mamonde",
    "BR000667": "MARICEEL",
    "BR000574": "Mary&amp;May",
    "BR000579": "MASIL",
    "BR000611": "Maxim",
    "BR000240": "Medicube",
    "BR000060": "Mediheal",
    "BR001378": "MediHeally",
    "BR000362": "MediPeel",
    "BR000210": "MeFactory",
    "BR000978": "Melixir",
    "BR000153": "Memebox",
    "BR001278": "MENOKIN",
    "BR000371": "MERBLISS",
    "BR000388": "Merzy",
    "BR001333": "Midha",
    "BR000507": "MIGUHARA",
    "BR000620": "Milk Touch",
    "BR001181": "Mimosu",
    "BR000287": "MineralBio",
    "BR000616": "MINIMUM",
    "BR000657": "MIRACLE M",
    "BR000031": "MiseEnScene",
    "BR000015": "Missha",
    "BR000872": "MIXSOON",
    "BR000144": "Mizon",
    "BR000468": "MLB",
    "BR001316": "MOEV",
    "BR001244": "MOIDA",
    "BR001294": "MOMMY CARE",
    "BR001332": "MONCLOS",
    "BR001173": "Moonseal",
    "BR000284": "moonshot",
    "BR000617": "Mude",
    "BR000515": "MULAWEAR",
    "BR001363": "MUMCHIT",
    "BR001279": "MUZIGAE MANSION",
    "BR001246": "My Fit",
    "BR000660": "MY1CART",
    "BR001102": "myFORMULA",
    "BR000320": "NACIFIC",
    "BR000263": "NAKE UP FACE",
    "BR000813": "NAMING",
    "BR000548": "NARD",
    "BR001203": "Narka",
    "BR000010": "Nature Republic",
    "BR001356": "NAUET",
    "BR001305": "NEAR",
    "BR001133": "Needly",
    "BR000205": "Neogen",
    "BR001290": "nesh",
    "BR001170": "NEULII",
    "BR001145": "NewTree",
    "BR000576": "NINE LESS",
    "BR001371": "No The Love",
    "BR001230": "NONOER",
    "BR000262": "nooni",
    "BR000891": "Numbuzin",
    "BR000647": "Nutri D-Day",
    "BR001344": "NUTSELINE",
    "BR001233": "OBgE",
    "BR001237": "ODDTYPE",
    "BR000815": "Ogi",
    "BR000304": "Olivarrier",
    "BR000538": "One-day's you",
    "BR000591": "Ongredients",
    "BR001199": "OOTD BEAUTY",
    "BR001348": "Orien",
    "BR000839": "P.CALM",
    "BR000322": "Pack age",
    "BR000207": "Paparecipe",
    "BR001160": "Parnell",
    "BR000033": "Peripera",
    "BR001360": "Pestlo",
    "BR000043": "Petitfee",
    "BR000618": "phykology",
    "BR000640": "PICKYWICKY",
    "BR001359": "Pleuvoir",
    "BR000286": "plu",
    "BR001337": "PODL",
    "BR001271": "Powerod",
    "BR000514": "PRAMY",
    "BR000476": "Preange",
    "BR001166": "PRESS SHOT",
    "BR000183": "Primera",
    "BR000597": "Pulmuone",
    "BR001225": "Pulmuone Garden Me",
    "BR001223": "PURCELL",
    "BR001004": "Puremay",
    "BR000524": "PURITO SEOUL",
    "BR000247": "Pyunkang yul",
    "BR001100": "RaNiq",
    "BR000602": "RAWEL",
    "BR001374": "RBOW",
    "BR000232": "RE:P",
    "BR000385": "Real Barrier",
    "BR001205": "Reboot",
    "BR001300": "REJURAN",
    "BR001268": "RETURNITY",
    "BR000317": "rom&amp;nd",
    "BR001287": "ROOTON",
    "BR000527": "Round Lab",
    "BR000329": "ROVECTIN",
    "BR000581": "SAFEAIR",
    "BR001174": "Saltysleep",
    "BR001011": "ScalpMed",
    "BR000475": "SCINIC",
    "BR001282": "seapuri",
    "BR000546": "Secret:X",
    "BR000178": "SecretKey",
    "BR001341": "SeohaeSol",
    "BR000568": "SERUMKIND",
    "BR000869": "SERY BOX",
    "BR001235": "SHAISHAISHAI",
    "BR000671": "SHAKE BABY",
    "BR001103": "simplyO",
    "BR000471": "sioris",
    "BR000996": "SKIN&amp;LAB",
    "BR000078": "Skin1004",
    "BR000017": "Skinfood",
    "BR001195": "Skinnylab",
    "BR000503": "SKINRx LAB",
    "BR001242": "slowpure",
    "BR000048": "SNP",
    "BR000223": "So natural",
    "BR001358": "SOLEP",
    "BR001186": "Someblossom",
    "BR000330": "SOMEBYMI",
    "BR000195": "SON & PARK",
    "BR001216": "SOOBLANC",
    "BR000443": "SOON MAMA",
    "BR000614": "SOON+",
    "BR000285": "SRB",
    "BR001194": "STUDIO 17",
    "BR000396": "Style by Aiahn",
    "BR001114": "STYLEKOREAN",
    "BR001283": "STYLEKOREAN BOX",
    "BR000007": "SU:M37˚",
    "BR001380": "SUELO",
    "BR000002": "Sulwhasoo",
    "BR001073": "Sungboon Editor",
    "BR000390": "Suntique",
    "BR000453": "SUR.MEDIC",
    "BR000569": "SUREBASE",
    "BR000212": "SWANICOCO",
    "BR001334": "TAMZ",
    "BR000578": "TEAZEN",
    "BR000441": "TENZERO",
    "BR000812": "TFIT",
    "BR001366": "The Creme Shop",
    "BR000008": "The History of Whoo",
    "BR000654": "THE LAB by blanc doux",
    "BR000440": "THE MASK SHOP",
    "BR000454": "THE ORDINARY",
    "BR000454": "The Plant Base (P'lab)",
    "BR001329": "The Purest Co",
    "BR000020": "the SAEM",
    "BR000006": "THEFACESHOP",
    "BR000280": "TIAM",
    "BR001301": "tiptoe",
    "BR000883": "TIRTIR",
    "BR000627": "TOCOBO",
    "BR001118": "Toi:L",
    "BR000011": "Tonymoly",
    "BR000016": "Too Cool For School",
    "BR000533": "Torriden",
    "BR000331": "TOSOWOONG",
    "BR000282": "touch in SOL",
    "BR000438": "Touch My body",
    "BR000519": "TOUN28",
    "BR001168": "treeannsea",
    "BR001243": "Treecell",
    "BR000270": "TROIAREUKE",
    "BR000534": "Twinkle Pop",
    "BR000854": "UNOVE",
    "BR001355": "UNRIPE",
    "BR000431": "URANG",
    "BR001259": "V Prove",
    "BR000064": "VDL",
    "BR001296": "Veganery",
    "BR001211": "Veganifect",
    "BR000392": "VELVIZO",
    "BR001210": "VERTTY",
    "BR001212": "VIDIVICI",
    "BR001328": "VIVELAB",
    "BR000470": "VIVLAS",
    "BR000307": "VT COSMETICS",
    "BR000200": "W.DRESSROOM",
    "BR000444": "WELLAGE",
    "BR001369": "Well-being Health Pharm",
    "BR000356": "WellDerma",
    "BR000289": "WHAMISA",
    "BR000258": "WonderBath",
    "BR000826": "Woori Nurungji",
    "BR000589": "Xpoiled",
    "BR000586": "YADAH",
    "BR000613": "YUNJAC",
    "BR001272": "z+piderm",
    "BR001289": "ZEROID",
    "BR000277": "ZYMOGEN"
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
        "Цена Discounted KRW", "Cena na site $", "Price", "Language", "Lower limit",
        "User group", "Особенности", "Старая цена KRW","status","category","procent"
    ])

    file_path = 'C:/Users/beeli/Downloads/parser_stas_final_2.xlsx'

    # --- Google Drive ---
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)
    folder_id = "10J-E4RcBJFfrdcqU_UAWask8BKTZ5Mw2"
    file_name = os.path.basename(file_path)

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
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000267", #G9
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000495", #GD11
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000518", #GILIM INTERNATIONAL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001185", #Ginger6
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001018", #Gleer
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001327", #glow
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000601", #Good Manner
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000168", #Goodal
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001229", #GOONGBE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001003", #GOUTTER
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001345", #GRABITY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001362", #GRAFEN
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000605", #Green Monster
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001134", #GRN Plus
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000416", #GROUNDPLAN
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001231", #GROWUS
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000577", #Gyeol Collagen
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001179", #HairPlus
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000988", #hakit
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=", #hanskin
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000026", #Hanyul
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000456", #Haruharu Wonder
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001379", #HAUA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000439", #hddn lab
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001373", #HealGrids
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000584", #Healing Bird
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000562", #Heart Percent
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000208", #Heimish
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000028", #HERA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000504", #Herbnote
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001276", #hetras
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001234", #HEVEBLUE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001165", #hince
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000022", #Holika Holika
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001075", #House of Hur
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001245", #HUECALM
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000273", #Huxley
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001069", #HYAAH
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000474", #Hydrogen
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000310", #HYGGEE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000540", #I DEW CARE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001013", #IBL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001196", #ID PLACOSMETICS
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000077", #ILLIYOON
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001164", #Ilso
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001288", #I'm meme
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000380", #I'm Sorry For My Skin
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000003", #Innisfree
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000005", #IOPE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000458", #ISNTREE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000535", #ISOI
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000023", #I's Skin
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000315", #IUNIK
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001169", #jaminkyung
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000672", #Jaunkyeol
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001071", #JAVIN DE SEOUL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000140", #Jayjun
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000623", #JelliFit
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000573", #Jenny House
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000341", #JMsolution
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000529", #J'S DERMA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000914", #Julie's Choice
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001330", #JULYME
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000296", #Jumiso
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000665", #Jungsaemmool
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001306", #JUST I AM
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001116", #JUST NATURE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000582", #KAHI
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000897", #KAINE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001104", #KAJA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000634", #KANU
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000366", #KEEP COOL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001207", #KEYTH
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001335", #KILLIT
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000459", #KIMJEONGMOON-ALOE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001324", #KINOHIMITSU
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000847", #Kirsh Blending
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001291", #Kitsui
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000604", #KIUKIMIUM
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000321", #KLAVUU
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001226", #KllureMY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001226", #KOELF
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001236", #KOPHER
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000599", #Korea Red Ginseng
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000585", #Kosette
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001191", #KSECRET
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001317", #KUNDAL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001005", #Kwailnara
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000206", #Labiotte
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001014", #LACTONIA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000358", #LAGOM
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001222", #LAKA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000437", #lalaChuu
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000949", #LALARECIPE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001175", #lalucell
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000004", #Laneige
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000019", #Leaders Insolution
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000186", #LEMONA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001357", #Libresse
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000913", #Lifepharm
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000490", #Lilybyred
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001343", #Lilyeve
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001183", #LINDSAY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000644", #LINGTEA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001322", #LINGTEA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000378", #LIZK
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000612", #LIZVIEW
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001178", #lleafill
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000603", #Lookas9
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000545", #Looks&Meii
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001101", #LYLA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000469", #ma:nyo
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001256", #Madeca 21
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000511", #MADECA DERMA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000989", #Makep:rem
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000651", #Mallingbooth
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000045", #Mamonde
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000667", #MARICEEL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000574", #Mary&May
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000579", #MASIL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000611", #Maxim
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000240", #Medicube
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000060", #Mediheal
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001378", #MediHeally
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000362", #MediPeel
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000210", #MeFactory
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000978", #Melixir
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000153", #Memebox
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001278", #MENOKIN
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000371", #MERBLISS
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000388", #Merzy
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001333", #Midha
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000507", #MIGUHARA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000620", #Milk Touch
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001181", #Mimosu
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000287", #MineralBio
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000616", #MINIMUM
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000657", #MIRACLE M
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000031", #MiseEnScene
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000015", #Missha
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000872", #MIXSOON
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000144", #Mizon
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000468", #MLB
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001316", #MOEV
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001244", #MOIDA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001294", #MOMMY CARE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001332", #MONCLOS
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001173", #Moonseal
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000284", #moonshot
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000617", #Mude
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000515", #MULAWEAR
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001363", #MUMCHIT
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001279", #MUZIGAE MANSION
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001246", #My Fit
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000660", #MY1CART
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001102", #myFORMULA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000320", #NACIFIC
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000263", #NAKE UP FACE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000813", #NAMING
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000548", #NARD
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001203", #Narka
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000010", #Nature Republic
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001356", #NAUET
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001305", #NEAR
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001133", #Needly
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000205", #Neogen
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001290", #nesh
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001170", #NEULII
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001145", #NewTree
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000576", #NINE LESS
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001371", #No The Love
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001230", #NONOER
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000262", #nooni
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000891", #Numbuzin
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000647", #Nutri D-Day
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001344", #NUTSELINE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001233", #OBgE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001237", #ODDTYPE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000815", #Ogi
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000304", #Olivarrier
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000538", #One-day's you
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000591", #Ongredients
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001199", #OOTD BEAUTY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001348", #Orien
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000839", #P.CALM
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000322", #Pack age
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000207", #Paparecipe
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001160", #Parnell
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000033", #Peripera
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001360", #Pestlo
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000043", #Petitfee
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000618", #phykology
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000640", #PICKYWICKY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001359", #Pleuvoir
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000286", #plu
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001337", #PODL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001271", #Powerod
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000514", #PRAMY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000476", #Preange
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001166", #PRESS SHOT
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000183", #Primera
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000597", #Pulmuone
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001225", #Pulmuone Garden Me
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001223", #PURCELL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001004", #Puremay
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000524", #PURITO SEOUL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000247", #Pyunkang yul
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001100", #RaNiq
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000602", #RAWEL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001374", #RBOW
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000232", #RE:P
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000385", #Real Barrier
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001205", #Reboot
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001300", #REJURAN
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001268", #RETURNITY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000317", #rom&nd
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001287", #ROOTON
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000527", #Round Lab
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000329", #ROVECTIN
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000581", #SAFEAIR
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001174", #Saltysleep
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001011", #ScalpMed
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000475", #SCINIC
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001282", #seapuri
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000546", #Secret:X
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000178", #SecretKey
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001341", #SeohaeSol
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000568", #SERUMKIND
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000869", #SERY BOX
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001235", #SHAISHAISHAI
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000671", #SHAKE BABY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001103", #simplyO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000471", #sioris
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000996", #SKIN&LAB
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000078", #Skin1004
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000017", #Skinfood
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001195", #Skinnylab
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000503", #SKINRx LAB
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001242", #slowpure
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000048", #SNP
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000223", #So natural
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001358", #SOLEP
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001186", #Someblossom
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000330", #SOMEBYMI
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000195", #SON & PARK
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001216", #SOOBLANC
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000443", #SOON MAMA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000614", #SOON+
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000285", #SRB
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001194", #STUDIO 17
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000396", #Style by Aiahn
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001114", #STYLEKOREAN
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001283", #STYLEKOREAN BOX
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000007", #SU:M37˚
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001380", #SUELO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000002", #Sulwhasoo
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001073", #Sungboon Editor
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000453", #Suntique
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000357", #SUR.MEDIC
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000569", #SUREBASE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000212", #SWANICOCO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001334", #TAMZ
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000578", #TEAZEN
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000441", #TENZERO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000812", #TFIT
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001366", #The Creme Shop
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000008", #The History of Whoo
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000654", #THE LAB by blanc doux
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000440", #THE MASK SHOP
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000454", #THE ORDINARY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000454", #The Plant Base (P'lab)
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001329", #The Purest Co
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000020", #the SAEM
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000006", #THEFACESHOP
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000280", #TIAM
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001301", #tiptoe
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000883", #TIRTIR
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000627", #TOCOBO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001118", #Toi:L
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000011", #Tonymoly
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000016", #Too Cool For School
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000533", #Torriden
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000331", #TOSOWOONG
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000282", #touch in SOL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000438", #Touch My body
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000519", #TOUN28
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001168", #treeannsea
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001243", #Treecell
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000270", #TROIAREUKE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000534", #Twinkle Pop
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000854", #UNOVE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001355", #UNRIPE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000431", #URANG
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001259", #V Prove
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000064", #VDL
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001296", #Veganery
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001211", #Veganifect
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000392", #VELVIZO
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001210", #VERTTY
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001212", #VIDIVICI
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001328", #VIVELAB
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000470", #VIVLAS
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000307", #VT COSMETICS
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000200", #W.DRESSROOM
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000444", #WELLAGE
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001369", #Well-being Health Pharm
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000356", #WellDerma
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000289", #WHAMISA
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000258", #WonderBath
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000826", #Woori Nurungji
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000589", #Xpoiled
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000586", #YADAH
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000613", #YUNJAC
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001272", #z+piderm
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR001289", #ZEROID
"https://stylekoreankbeautywholesale.com/Product/BrandProduct?brand_cd=BR000277", #ZYMOGEN
    ]

    for brand_url in brand_urls:
        brand_name = extract_brand_name(brand_url)
        print(f"Scraping products for brand: {brand_name}")

        driver.get(brand_url)
        handle_alert(driver)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "album")))

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
                    name_el = card.find("span", class_="productTxt")
                    name = name_el.get_text(strip=True) if name_el else ""

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

                    # --- STATUS ---
                    brand_for_status = brand_name_clean if brand_name_clean else brand_name
                    status_value = f"Бренд///{brand_for_status[:1].upper() if brand_for_status else 'X'}///{brand_for_status}"
                    STATUS = "A"

                    # --- CLEAN CODES ---
                    item_code_clean = re.sub(r'\s+', '', item_code) if item_code else ""
                    product_code_clean = re.sub(r'\s+', '', Product_code) if Product_code else ""

                    # --- APPEND ROW ---
                    ws.append([
                        img_src,
                        brand_name_clean if brand_name_clean else brand_name,
                        name,
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
                        procent
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
                query = f"'{folder_id}' in parents and trashed=false and title='{file_name}'"
                file_list = drive.ListFile({'q': query}).GetList()
                if file_list:
                    file_list[0].Delete()
                    print(f"Старый файл '{file_name}' удалён с Google Drive")

                file_drive = drive.CreateFile({'title': file_name, 'parents':[{'id': folder_id}]})
                file_drive.SetContentFile(file_path)
                file_drive.Upload()
                print(f"Файл '{file_name}' успешно обновлён на Google Drive после страницы {page_num}")
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
