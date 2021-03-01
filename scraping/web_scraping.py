def create_soup(url):

    #  라이브러리
    import requests
    from bs4 import BeautifulSoup

    #  User_Agnet
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36"}

    #  리퀘스트
    url = f"{url}"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "lxml")

    return soup


def create_selenium(url):

    #  라이브러리

    from selenium import webdriver

    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.implicitly_wait(3)

    # 페이지 오픈
    url = f"{url}"
    driver = driver.get(url)
    return driver

#   와이 함수로 셀레니움을 실행하면 저절로 창이 닫히는지?

def create_headless_selenium(url):

    # 라이브러리

    from selenium import webdriver

    # headless
    options = webdriver.ChromeOptions()
    options.headless = True
    options.add_argument("window-size=1920x1080")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36")

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.implicitly_wait(3)

    url = f"{url}"
    driver.get(url)
    return driver.get(url)


def create_csv(title, *header):

    #  라이브러리
    import csv

    #  csv 파일 생성
    csv_file = open(f"{title}.csv", "a", encoding='utf-8-sig', newline="")
    csv_writer = csv.writer(csv_file)

    # csv 헤더
    csv_header = []
    for line in header:
        csv_header.append(line)

    csv_writer.writerow(csv_header)


# def create_xlsx(ws_title, wb_title, *header):
#     #  라이브러리
#     from openpyxl import Workbook
#
#     #  워크북 생성
#     wb = Workbook()
#     ws = wb.active  #  현재 열려있는 페이지
#     ws.title = f"{ws_title}"
#     wb.create_sheet()  # 활성화 된 시트 옆에 새로운 시트를 만들면서 이름을 부여함
#
#     #  엑셀 헤더
#     csv_header = []
#     for line in header:
#         csv_header.append(line)


