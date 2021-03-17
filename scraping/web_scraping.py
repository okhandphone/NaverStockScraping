def create_soup(url):
    import requests
    from bs4 import BeautifulSoup

    #  User_Agnet
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36"}

    url = f"{url}"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "lxml")
    return soup


'''
selenium 이렇게 불러오니까 속도 개느림
'''
def create_selenium(url):
    from selenium import webdriver

    options = webdriver.ChromeOptions()
    options.headless=True
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36")
    driver = webdriver.Chrome(options=options)

    url = f"{url}"
    driver.get(url)

    return driver

#   와이 함수로 셀레니움을 실행하면 저절로 창이 닫히는지?


def create_csv(title, *header):
    import csv

    #  csv 파일 생성
    csv_file = open(f"{title}.csv", "a", encoding='utf-8-sig', newline="")
    csv_writer = csv.writer(csv_file)

    # 헤더
    header_list = []
    for line in header:
        header_list.append(line)

    csv_writer.writerow(header_list)


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

# 페이지 스크롤
# 페이지 살 짝 내려서 댓글 나타나게 하기

# driver.find_element_by_css_selector('html').send_keys(Keys.PAGE_DOWN)


# # 코드잇
# last_height = driver.execute_script("return document.body.scrollHeight")
# while True:
#     driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#     time.sleep(1)
#
#     new_height = driver.execute_script("return document.body.scrollHeight")
#     if new_height == last_height:
#         break
#     last_height = new_height
#
# 나도코딩
# prev_height = driver.execute_script("return document.body.scrollHeight")
# print('스크롤시작')
# # 반복 수행
# while True:
#     # 스크롤을 가장 아래로 내림
#     driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
#     # 페이지 로딩 대기
#     time.sleep(1)
#     # 현재 문서 높이를 가져와서 저장
#     curr_height = driver.execute_script("return document.body.scrollHeight")
#     if curr_height == prev_height:
#         break
#     prev_height = curr_height
#     print("스크롤 반복")