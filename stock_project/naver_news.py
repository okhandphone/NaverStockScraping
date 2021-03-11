import requests
from bs4 import BeautifulSoup
import os
import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image


# 뉴스 저장 경로 및 파일 만들기
news_path = "stock_data\\News"
if not os.path.exists(news_path + "\\데일리 뉴스.xlsx"):
    os.mkdir(news_path)

    wb = Workbook()
    ws = wb.active
    ws.title = "데일리 뉴스"
    ws.append(['링크', '키워드', '날짜', '타이틀', '요약'])

    # 속성 {열번호: 너비}
    attrs = {"B": "12", "C": "11", "D": "49"}
    for key, value in attrs.items():
        ws.column_dimensions[f"{key}"].width = f"{value}"

    # for i in range(1, ws.max_column):
    #     ws[f"A{i}"].font.copy(bold=True)
    # ws.rows(1).alignment = Alignment(ho)
    wb.save(news_path + "\\데일리 뉴스.xlsx")
    print("새 엑셀 파일 생성")

# 이미 디렉토리가 있다면 엑셀 파일 로드
else:
    wb = load_workbook(news_path + "\\데일리 뉴스.xlsx")
    ws = wb.active
    print("데일리뉴스 엑셀파일 로드")

while True:
    keyword = input("키워드 입력 : ")
    # keyword = ['sk', '현대차', '현대모비스', '기아차', '미국 증시', '연준', 'FOMC', '끝']

    # 끝이라고 입력하면 프로그램 종료
    if keyword == "끝":
        break

    # # 이미지 폴더 만들기
    # img_path = f"stock_data\\News\\image"
    # if not os.path.exists(img_path):
    #     os.mkdir(img_path)

    # 오늘날짜 함께 입력
    today = datetime.date.today()
    print(today)
    # 뉴스 페이지 '최신순' 첫번재 페이지
    url = f"https://search.naver.com/search.naver?where=news&query={keyword}&sm=tab_srt&sort=1&photo=0&field=0&reporter_article=&pd=0&ds=&de=&docid=&nso=so%3Add%2Cp%3Aall%2Ca%3Aall&mynews=0&refresh_start=0&related=0"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.72 Safari/537.36 Edg/89.0.774.45"}
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, "html.parser")

    t = soup.select('.news_area > a')
    img = soup.select('.news_wrap .dsc_thumb img')

    for i in range(len(t)):
        try:
            # 타이틀 및 링크
            title = t[i].text
            link = t[i]['href']

            ws.append([link, keyword, today, title])
            wb.save(news_path + "\\데일리 뉴스.xlsx")

            # # 내용 받아오기
            # news_res = requests.get(link)
            # news_soup = BeautifulSoup(news_res, "html.parser")
            # content = news_soup.select_one("subTitle_s2 br").text.strip()

            # # 이미지도 함께 다운 / 입력
            # img_url = img[i]['src']
            # response = requests.get(img_url)
            # response.raise_for_status()
            #
            # with open(img_path + f"\\{keyword}_{i + 1}.png", "wb") as f:
            #     f.write(response.content)
            #
            # img_for_excel = Image(img_path + f"\\{keyword}_{i + 1}.png")
            # insert_img = ws.add_image(img_for_excel, f"A{ws.max_row}")
            # ws.append([insert_img, link, keyword, today, title])
        except:
            pass