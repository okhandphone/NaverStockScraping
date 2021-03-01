# This Python file uses the following encoding: utf-8
import os, sys

# # 번갈아 가면서 작동 ㄴㄴ
#
#
# def get_real_time_value(mk_code_dict):
#     import time
#     import csv
#     from selenium import webdriver
#     from selenium.webdriver.common.by import By
#     from selenium.webdriver.support.ui import WebDriverWait
#     from selenium.webdriver.support import expected_conditions as EC
#     from bs4 import BeautifulSoup
#
#
#     # CSV
#     csv_file = open('종목 실시간 밸류에이션.csv', 'a', encoding='utf-8-sig', newline='')
#     csv_writer = csv.writer(csv_file)
#     csv_writer.writerow(['종목코드', '종목명', '업종코드', '업종명', '시가총액', 'PER', 'ROE', 'ROA', 'PBR', '유보율'])
#
#     #  headless
#     options = webdriver.ChromeOptions()
#     options.headless = True
#     options.add_argument("window-size=1920x1080")
#     options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36")
#     driver = webdriver.Chrome(options=options)
#     driver.maximize_window()
#     print("Get real time value start:")
#     print(time.strftime("%Y-%m-%d %H:%M:%S"))
#     #  mk_code_dict를 파라미터로
#
#     time.sleep(3)
#     print(f"Start {value} Category")
#     driver.get(f'https://finance.naver.com/sise/sise_group_detail.nhn?type=upjong&no={value}')
#     wait = WebDriverWait(driver, 3)
#
#     # 기존 옵션 제거
#     wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option1'))).click()  # 거래량
#     wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option2'))).click()  # 매수호가
#     wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option3'))).click()  # 거래대금
#     wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option8'))).click()  # 매도호가
#     wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option9'))).click()  # 전일 거래량
#
#     # 원하는 옵션 클릭
#     wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option4'))).click()  # 시가총액
#     wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option6'))).click()  # PER
#     wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option12'))).click()  # ROE
#     wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option18'))).click()  # ROA
#     wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option24'))).click()  # PBR
#     wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option27'))).click()  # 유보율
#
#     # 옵션 적용 클릭
#     wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.item_btn > a'))).click()  # 적용하기
#
#     for key, value in mk_code_dict.items():
#         try:
#             time.sleep(3)
#             print(f"Start {value} Category")
#             driver.get(f'https://finance.naver.com/sise/sise_group_detail.nhn?type=upjong&no={value}')
#
#             driver.save_screenshot(f"screenshot{key}.png")
#             # 실시간 기업 밸류에이션 스크래핑
#             soup = BeautifulSoup(driver.page_source, "html.parser")
#             table = soup.select('#contentarea > div.box_type_l')[1].select('tbody > tr')[:-2]
#             for line in table:
#                 share_name = line.td.text
#                 share_code = line.td.a['href'].replace("/item/main.nhn?code=", "")
#                 info_list = [share_code, share_name, value, key]
#
#                 data = line.select('td')[4:-1]
#                 for num in data:
#                     num = num.text.replace(',', '')
#                     info_list.append(num)
#                 csv_writer.writerow(info_list)
#               # ws.append(info_list)
#
#         except:
#             print(f"err code: {value}")
#             pass
#     # wb.save("종목 실시간 벨류에이션.xlsx")
#     driver.quit()
#     csv_file.close()
#     print("RealTimeValue Completed")
#     print(time.strftime("%Y-%m-%d %H:%M:%S"))
#
# mk_code_dict = {'증권': '12', '건설': '42', '손해보험': '190', '디스플레이장비및부품': '199', '섬유,의류,신발,호화품': '134'}
# get_real_time_value(mk_code_dict)

# data = [77,
# 14745,
# 6380,
# 6390,
# 3486914,
# 22755,
# 5129739,
# 20200,
# 20250,
# 7343076,
# 155598,
# 6702288]
# #
# # for num in data:
# #     print(num)
#
# [print(i) for i in data]

# get_share_code()의 리턴값 코드 딕셔너리를 받아서 연별, 분기별 실적 정리
# code_dict = {'016610': {'Share Name': 'DB금융투자 ', 'Market Code': '12', 'Market Name': '증권'}, '375500': {'Share Name': 'DL이앤씨 ', 'Market Code': '42', 'Market Name': '건설'}, '112190': {'Share Name': 'DB손해보험 ', 'Market Code': '190', 'Market Name': '손해보험'},'068790': {'Share Name': 'DMS *', 'Market Code': '199', 'Market Name': '디스플레이장비및부품'}, '001530': {'Share Name': 'DI동일 ', 'Market Code': '134', 'Market Name': '섬유,의류,신발,호화품'}}
# import time
# from bs4 import BeautifulSoup
# from pandas import DataFrame
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
#
# # headless webdriver
# options = webdriver.ChromeOptions()
# options.headless = True
# options.add_argument("window-size=1920x1080")
# options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36")
# driver = webdriver.Chrome(options=options)
# driver.maximize_window()
# driver.implicitly_wait(3)
#
# # header = ",Share Code,Share Name,Market Code,Market Name,매출액,영업이익,영업이익(발표기준),세전계속사업이익,당기순이익,  당기순이익(지배),  당기순이익(비지배),자산총계,부채총계,자본총계,  자본총계(지배),  자본총계(비지배),자본금,영업활동현금흐름,투자활동현금흐름,재무활동현금흐름,CAPEX,FCF,이자발생부채,영업이익률,순이익률,ROE(%),ROA(%),부채비율,자본유보율,EPS(원),PER(배),BPS(원),PBR(배),현금DPS(원),현금배당수익률,현금배당성향(%),발행주식수(보통주)".split(',')
#
# print("Get Year Financial Info Start : ")
# print(time.strftime("%Y-%m-%d %H:%M:%S"))
# code_dict = {'016610': {'Share Name': 'DB금융투자 ', 'Market Code': '12', 'Market Name': '증권'}, '005830': {'Share Name': 'DB손해보험 ', 'Market Code': '190', 'Market Name': '손해보험'}, '112190': {'Share Name': 'KC산업 ', 'Market Code': '191', 'Market Name': '상사'}, '000990': {'Share Name': 'DB하이텍 ', 'Market Code': '202', 'Market Name': '반도체와반도체장비'}, '000995': {'Share Name': 'DB하이텍1우 ', 'Market Code': '202', 'Market Name': '반도체와반도체장비'}, '139130': {'Share Name': 'DGB금융지주 ', 'Market Code': '20', 'Market Name': '은행'}, '001530': {'Share Name': 'DI동일 ', 'Market Code': '134', 'Market Name': '섬유,의류,신발,호화품'}, '000210': {'Share Name': 'DL ', 'Market Code': '42', 'Market Name': '건설'}, '000215': {'Share Name': 'DL우 ', 'Market Code': '42', 'Market Name': '건설'}, '375500': {'Share Name': 'DL이앤씨 ', 'Market Code': '42', 'Market Name': '건설'}, '068790': {'Share Name': 'DMS *', 'Market Code': '199', 'Market Name': '디스플레이장비및부품'}, '112190': {'Share Name': 'DRB동일 ', 'Market Code': '36', 'Market Name': '자동차부품'}}
#
# for key, value in code_dict.items():
#     try:
#         time.sleep(3)
#         print(f"Start: {key}")
#         driver.get(f'https://finance.naver.com/item/coinfo.nhn?code={key}')
#         wait = WebDriverWait(driver, 3)
#
#         # iframe 접근
#         driver.switch_to.frame('coinfo_cp')
#         time.sleep(1.5)
#
#         # 연간 분기 돌아가면서 정보 받아오기
#         period_dict = {'cns_Tab21': '연간', 'cns_Tab22': '분기'}
#         for p_key, p_val in period_dict.items():
#             wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'#{p_key}'))).click()
#             time.sleep(1)
#
#             # iframe 에서 html정보 추출
#             soup = BeautifulSoup(driver.page_source, "html.parser")
#
#             # 기간 정보
#             period_data = soup.select('table.gHead01')[3].select('thead > tr > th')[2:7]
#             period = [line.text.strip()[:7] for line in period_data]
#
#             # financial_dict에 코드정보 입력
#             financial_dict = {'Share Code': f'{str(key)}'}
#             financial_dict.update(value)  # code_dict의 정보를 추가
#
#             # 재무정보
#             financial_table = soup.select('table.gHead01')[3].select('tbody > tr')
#             for line in financial_table:
#                 name = line.select_one('th').text  # 재무 항목명
#                 data = line.select('td')[:5]  # 재무 데이터 : 앞 다섯 기간만, 컨센 제외
#                 financial_dict.setdefault(name, [num.text.replace(",", "") for num in data])
#
#             # financial_dict를 데이터 프레임으로
#             # 컬럼 = list(financial_dict.keys()) = name, 인덱스 = 기간정보
#             financial_df = DataFrame(financial_dict, columns=list(financial_dict.keys()), index=period)
#             financial_df.to_csv(f"{p_val}실적_raw.csv", mode="a", header=False, encoding='utf-8-sig')
#     except:
#         pass
#         print(f'pass: {key}')
# print("Financial Info Done")
# print(time.strftime("%Y-%m-%d %H:%M:%S"))
# driver.quit()


'''def get_quarter_financial_info(code_dict):

    from bs4 import BeautifulSoup
    from pandas import DataFrame
    from selenium import webdriver
    import time

    # headless
    options = webdriver.ChromeOptions()
    options.headless = True
    options.add_argument("window-size=1920x1080")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36")

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.implicitly_wait(3)

    # get_share_code()의 리턴값 코드 딕셔너리를 받아서 분기 정보를 정리

    quarter_financial_dict = {}
    for key in code_dict.keys():
        driver.get(f'https://finance.naver.com/item/coinfo.nhn?code={key}')

        # iframe 에서 html정보 추출
        driver.switch_to.frame('coinfo_cp')

        # 분기 정보
        time.sleep(1.0)
        driver.find_element_by_id('cns_Tab22').click()
        time.sleep(0.5)
        soup = BeautifulSoup(driver.page_source, "html.parser")

        # 기간 정보
        period_data = soup.select('table.gHead01')[3].select('thead > tr > th')[2:7]
        period_quarter = []
        for line in period_data:
            date = line.text.strip()[:7]
            period_quarter.append(date)
        # print(period_quarter)

        # 재무정보
        financial_table = soup.select('table.gHead01')[3].select('tbody > tr')
        # code_dict의 정보를 추가하여 financial _dict 생성
        # quarter_financial_dict = {"Share Code": f'{key}',
        #                           'Share Name': code_dict[f'{key}']['Share Name'],
        #                           'Market Code': code_dict[f'{key}']['Market Code'],
        #                           'Market Name': code_dict[f'{key}']['Market Name']}
        for line in financial_table:
            name = line.select_one('th').text
            data = line.select('td')[:5]
            # 숫자는 실수로 변경
            for num in data:
                num = num.text.replace(",", "")
                if num == 'N/A':
                    num = ''
                elif num == '':
                    num = ''
                else:
                    num = float(num)
                # financial_dict : key = name, value = data num list
                quarter_financial_dict.setdefault(name, []).append(num)

        # financial_dict를 데이터 프레임으로 : 컬럼 = financial_dict.keys() = name, 로우 인덱스 = 기간정보
        quarter_financial_df = DataFrame(quarter_financial_dict, columns=list(quarter_financial_dict.keys()), index=period_quarter)
        # quarter_financial_df.index.name = 'Period'

    driver.quit()
    # return quarter_financial_dict

# print(get_quarter_financial_info(code_dict))



# 3. 연별 실적 스크래핑

def get_year_financial_info(code_dict):
    from bs4 import BeautifulSoup
    from pandas import DataFrame
    from selenium import webdriver
    import time

    # headless
    options = webdriver.ChromeOptions()
    options.headless = True
    options.add_argument("window-size=1920x1080")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36")

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.implicitly_wait(3)


    print("Get Year Financial Info Start")
    # get_share_code()의 리턴값 코드 딕셔너리를 받아서 연별실적 정리
    for key, value in code_dict.items():
        driver.get(f'https://finance.naver.com/item/coinfo.nhn?code={key}')

        # iframe 에서 html정보 추출
        driver.switch_to.frame('coinfo_cp')

        # 연간 정보 접근
        time.sleep(1)
        driver.find_element_by_id('cns_Tab21').click()
        time.sleep(0.5)
        soup = BeautifulSoup(driver.page_source, "html.parser")

        # 기간 정보
        period_data = soup.select('table.gHead01')[3].select('thead > tr > th')[2:7]
        period_year = []
        for line in period_data:
            date = line.text.strip()[:4]
            period_year.append(date)

        # 재무정보
        financial_table = soup.select('table.gHead01')[3].select('tbody > tr')

        # financial _dict에 기간별 재무정보 입력
        year_financial_dict = {'Share Code': f'{str(key)}'}
        year_financial_dict.update(value)  # code_dict의 정보를 추가

        for line in financial_table:
            name = line.select_one('th').text  # 재무 항목명
            data = line.select('td')[:5]  # 재무 데이터
            for num in data:
                num = num.text.replace(",", "")
                # financial_dict 구성 : key = name, value = num list
                year_financial_dict.setdefault(name, []).append(num)

        # financial_dict를 데이터 프레임으로 : 컬럼 = financial_dict.keys() = name, 로우 인덱스 = 기간정보
        year_financial_df = DataFrame(year_financial_dict, index=period_year)
        year_financial_df.to_csv("연별실적.csv", mode="a", header=False, encoding='utf-8-sig')
        # year_financial_df.index.name = 'Period'
        # print(year_financial_df)
    driver.quit()
    print("Fianancial Info Done")

    # return year_financial_dict


# from kyu_def.stock_scraping_master import get_etf_info
# etf_code = ['069500', '367380', '341850', '376250', '245350']

import time
import datetime
import csv
from kyu_def.web_scraping import create_soup
from selenium import webdriver
from bs4 import BeautifulSoup

# # CSV
# csv_file = open(f'etf_{datetime.date.today()}.csv', 'a', encoding='utf-8-sig', newline='')
csv_writer = csv.writer(csv_file)
csv_writer.writerow(['ETF이름', '링크', 'ETF코드', '운용사', '수수료', '시가총액(억 원)', '수익률', '', '', '', '구성종목 TOP 10'])
csv_writer.writerow(['', '', '', '', '', '', '1개월', '3개월', '6개월', '1년', '1위', '', '2위', '', '3위', '', '4위', '', '5위', '', '6위', '','7위', '', '8위', '', '9위', '', '10위', ''])

# 헤드리스 셀레니움
options = webdriver.ChromeOptions()
options.headless = True
options.add_argument("window-size=1920x1080")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36")
driver = webdriver.Chrome(options=options)
driver.maximize_window()
driver.implicitly_wait(3)

# print(f"Get ETF info :")
# print(time.strftime("%Y-%m-%d %H:%M:%S"))


url = 'https://finance.naver.com/item/coinfo.nhn?code=52420'
driver.get(url)
soup = create_soup(url)
print(soup)
# 이름, 링크, 운용사, 수수료
etf_name = soup.select('#middle')


etf_link = f'https://finance.naver.com/item/coinfo.nhn?code='
company = soup.select_one('#tab_con1').select_one('div:nth-child(4)').select('td')[1].text
commission = soup.select_one('#tab_con1').select_one('div:nth-child(4)').select('td')[0].em.text.replace('%', '')
each_etf_info = [etf_name, etf_link, str(code), company, commission] # 1차 리스트 저장

# iframe으로 이동
driver.switch_to.frame('coinfo_cp')
time.sleep(1)
soup_iframe = BeautifulSoup(driver.page_source, "lxml")

# 시가총액
capital = soup_iframe.select_one('#status_grid_body').select('tr')[5].td.text.replace(',', '')[:-2]
each_etf_info.append(capital)

# 수익률
earnings = soup_iframe.select_one('#status_grid_body').select('tr')[-1].td.select('span')
for line in earnings:
    earning = line.text.replace('%', '').replace('+', '')
    each_etf_info.append(earning)

# 구성종목
top_10 = soup_iframe.select_one('#CU_grid_body').select('tr')[:10]
for line in top_10:
    sh_name = line.select('td')[0].text
    percent = line.select('td')[2].text
    if percent == '-':
        percent = ''
    else:
        percent = percent
    # 리스트 진짜 이방법밖에 없는거 실화..? 더 찾아보기
    each_etf_info.append(sh_name)  # 3차 리스트 저장
    each_etf_info.append(percent)
# csv에 입력
csv_writer.writerow(each_etf_info)

driver.quit()
print(f"ETF info completed")
print(time.strftime("%Y-%m-%d %H:%M:%S"))
# for code in etf_code:
#     time.sleep(5)
#     print(f"Start : {code}")
#     url = f'https://finance.naver.com/item/coinfo.nhn?code={code}'
#     driver.get(url)
#     soup = create_soup(url)
#
#     # 이름, 링크, 운용사, 수수료
#     etf_name = soup.select_one('#middle > div.h_company > div.wrap_company > h2 > a').text
#     print(etf_name)
#     etf_link = f'https://finance.naver.com/item/coinfo.nhn?code={code}'
#     company = soup.select_one('#tab_con1').select_one('div:nth-child(4)').select('td')[1].text
#     commission = soup.select_one('#tab_con1').select_one('div:nth-child(4)').select('td')[0].em.text.replace('%', '')
#     each_etf_info = [etf_name, etf_link, str(code), company, commission] # 1차 리스트 저장
#
#     # iframe으로 이동
#     driver.switch_to.frame('coinfo_cp')
#     time.sleep(1)
#     soup_iframe = BeautifulSoup(driver.page_source, "lxml")
#
#     # 시가총액
#     capital = soup_iframe.select_one('#status_grid_body').select('tr')[5].td.text.replace(',', '')[:-2]
#     each_etf_info.append(capital)
#
#     # 수익률
#     earnings = soup_iframe.select_one('#status_grid_body').select('tr')[-1].td.select('span')
#     for line in earnings:
#         earning = line.text.replace('%', '').replace('+', '')
#         each_etf_info.append(earning)
#
#     # 구성종목
#     top_10 = soup_iframe.select_one('#CU_grid_body').select('tr')[:10]
#     for line in top_10:
#         sh_name = line.select('td')[0].text
#         percent = line.select('td')[2].text
#         if percent == '-':
#             percent = ''
#         else:
#             percent = percent
#         # 리스트 진짜 이방법밖에 없는거 실화..? 더 찾아보기
#         each_etf_info.append(sh_name)  # 3차 리스트 저장
#         each_etf_info.append(percent)
#     # csv에 입력
#     csv_writer.writerow(each_etf_info)
#
# driver.quit()
# print(f"ETF info completed")
# print(time.strftime("%Y-%m-%d %H:%M:%S"))'''