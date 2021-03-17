
'''
엑셀 각종 서식 모음 파일
'''


import os
import csv
import time
import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from scraping.stock_scraping_master import get_naver_market_code
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Color, Side
import pandas as pd

'''
데이터 프레임으로 csv 파일 엑셀에 붙여넣기
'''
file_name = dir_path + "/realtime_stock_value.xlsx"
df = pd.read_csv(dir_path + f'/realtime_stock_value_{datetime.date.today()}.csv')

# pd.ExcelWriter() : 기존 파일 지워짐
if not os.path.exists(file_name):
    with pd.ExcelWriter(file_name, mode='w', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=f"{datetime.date.today()}",index=False)
else:
    with pd.ExcelWriter(file_name, mode='a', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=f"{datetime.date.today()}",index=False)

print("df to 엑셀")


'''
엑셀 각종 서식 모음
'''

# 셀서식
align_center = Alignment(horizontal="center", vertical="center")
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

# 글꼴 서식
font_bold = Font(name="맑은 고딕", size=11, bold=True)
font_red = Font(name="맑은 고딕", size=11, color="FE2E2E")
font_blue = Font(name="맑은 고딕", size=11, color="1E88E5")

# color_fill
green_fill = PatternFill(patternType="solid", fgColor=Color("E9F7EF"))
orange_fill = PatternFill(patternType="solid", fgColor=Color("FEF5E7"))

light_pink_fill = PatternFill(patternType="solid", fgColor=Color("FDEDEC"))
pink_fill = PatternFill(patternType="solid", fgColor=Color("F5B7B1"))
dark_pink_fill = PatternFill(patternType="solid", fgColor=Color("EC7063"))
red_fill = PatternFill(patternType="solid", fgColor=Color("E74C3C"))

light_blue_fill = PatternFill(patternType="solid", fgColor=Color("D6EAF8"))
blue_fill = PatternFill(patternType="solid", fgColor=Color("AED6F1"))
dark_blue_fill = PatternFill(patternType="solid", fgColor=Color("5DADE2"))
navy_fill = PatternFill(patternType="solid", fgColor=Color("2E86C1"))

# 워크북 불러오기
wb = load_workbook(file_name)
ws = wb[f"{datetime.date.today()}"]

# 컬럼너비 (종목명, 업종명, 상승종목 수, 하락종목 수)
width_dict = {"C": 21, "E": 19, "O": 12, "P": 12}
for key, value in width_dict:
    ws.column_dimensions[f"{key}"].width = value

# 헤더 서식
for row in ws["A1:M1"]:
    for cell in row:
        cell.font = font_bold
        cell.alignment = align_center
        cell.fill = orange_fill
        cell.border = thin_border

# 등락률 글자색
count_inc = 0 # 상승 종목 갯수
count_dec = 0 # 하락 종목 갯수
for col in ws[f"F2:F{ws.max_row}"]:
    for cell in col:
        if cell.value > 0:
            cell.font = font_red
            count_inc += 1
        else:
            cell.font = font_blue
            count_dec += 1

# 하이퍼링크
for i in range(2, ws.max_row):
    ws[f"C{i}"].hyperlink = ws[f"A{i}"].value
    print(ws[f"A{i}"].value)

# 상승 하락 종목 셀 서식
ws["O1"].value = "상승 종목 수"
ws["P1"].value = "하락 종목 수"
ws["O1"].fill = dark_pink_fill
ws["P1"].fill = dark_blue_fill
ws["O2"].value = count_inc
ws["P2"].value = count_dec

for row in ws[f"O1:P1"]:
    for cell in row:
        cell.font = font_bold

for row in ws[f"O1:P2"]:
    for cell in row:
        cell.border = thin_border
        cell.alignment = align_center

print(count_inc, count_dec)

# 전체 보더
for row in ws[f"A2:M{ws.max_row}"]:
    for cell in row:
        cell.border = thin_border

# 오토필터
ws.auto_filter.ref = ws.dimensions
ws.freeze_panes = 'A2'

# 컬럼 숨기기 (URL)
ws.column_dimensions["A"].hidden= True

wb.save(file_name)
print("엑셀 서식 완료")