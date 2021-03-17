# This Python file uses the following encoding: utf-8
import os, sys


#  User-Agent
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36"}


'''
1. 업종 종목 코드
'''

# 0. 한국거래소 종목, 업종 코드
# 엑셀 파일 아닌 홈페이지에서 바로 받는 법 찾아보기
def get_share_code():
    # openpyxl로 엑셀 파일 읽기
    from openpyxl import load_workbook  # 파일 불러오기
    wb = load_workbook("stock_data\\krxcode.xlsx")  # 불러올 파일명 넣기
    ws = wb.active

    print("get Share Code Dict")
    code_list = []
    code_dict = {}
    for x in range(2, ws.max_row + 1):
        code = [ws.cell(row=x, column=y).value for y in range(1, ws.max_column + 1)]
        code_list.append(code)
        code_dict.setdefault(f"{code[0]}", {"Share Name": code[1], "Market Code": str(code[2]), "Market Name": code[3]})

    print("get Share Code Done")
    return code_dict

# 1. 네이버에서 업종코드 받기
def get_naver_market_code():
    from scraping.web_scraping import create_soup
    soup = create_soup('https://finance.naver.com/sise/sise_group.nhn?type=upjong')
    table = soup.select('table.type_1 > tr > td > a')
    mk_code_dict = {}
    for line in table:
        market_name = line.text
        market_code = line['href'].replace("/sise/sise_group_detail.nhn?type=upjong&no=", "")
        mk_code_dict[market_name] = market_code
    del mk_code_dict['기타']
    print(mk_code_dict)
    print('mk_code_dict completed')
    # mk_code_dict를 리턴 {업종명: 업종코드}
    return mk_code_dict

# 2. 네이버 업종코드를 통해 종목 코드 받기
def get_share_code_from_naver(mk_code_dict):
    from scraping.web_scraping import create_soup
    from openpyxl import Workbook
    import time

    wb = Workbook()
    ws = wb.active
    ws.title = "네이버 업종 종목코드"
    ws.append(['종목코드', '종목명', '업종코드', '업종명'])

    for key, value in mk_code_dict.items():
        print(f"Start {value} Category")
        # time.sleep(2)
        url = f'https://finance.naver.com/sise/sise_group_detail.nhn?type=upjong&no={value}'
        soup = create_soup(url)
        table = soup.select('#contentarea > div.box_type_l')[1].select('tbody > tr > td.name')
        for line in table:
            share_name = line.text
            share_code = line.a['href'].replace("/item/main.nhn?code=", "")
            # 마켓과 종목 코드를 엑셀파일로 저장
            ws.append([share_code, share_name, value, key])
    wb.save("stock_data\\mk_sh_code_naver.xlsx")
    print("Total Code file Completed")

# 3. 엑셀에서 업종코드 딕셔너리로 가져오기
def get_code_from_excel():
    from openpyxl import load_workbook  # 파일 불러오기
    # 상대경로로 수정해야함
    # wb = load_workbook("C:\\Users\\MING\\PycharmProjects\\web_automation\\stock_project\\mk_sh_code_naver.xlsx")  # 불러올 파일명 넣기
    wb = load_workbook("stock_data\\mk_sh_code_naver.xlsx")
    ws = wb.active

    code_dict = {}
    for x in range(2, ws.max_row + 1):
        code = [ws.cell(row=x, column=y).value for y in range(1, ws.max_column + 1)]
        code_dict.setdefault(f"{code[0]}", {"Share Name": code[1], "Market Code": str(code[2]), "Market Name": code[3]})
    print(code_dict)
    print("code_dict completed")
    # 종합 코드딕트 생성 {종콕코드 : {"Share Name": 종목명, "Market Code": 업종코드, "Market Name": 업종명}
    return code_dict



'''
2. 네이버 테마 정보
'''

#  4. 네이버 테마 코드 받기
def get_naver_theme_code():
    import time
    from scraping.web_scraping import create_soup
    theme_code_list = [] # 테마 코드 리스트로 저장
    print("Start get share code")
    for page in range(1, 7):
        print(f"On page : {page}")
        # 네이버 금융 테마별 시세
        # time.sleep(3)
        url = f"https://finance.naver.com/sise/theme.nhn?&page={page}"
        soup = create_soup(url)

        theme_data = soup.select_one("table.type_1").select("td.col_type1")
        for line in theme_data:
            theme_code = line.select_one("a")["href"].replace("/sise/sise_group_detail.nhn?type=theme&no=", "")
            theme_code_list.append(theme_code)
    print("Theme code done")
    # 테마 코드 리스트를 리턴
    return theme_code_list


#  5. 네이버 테마별 테마명, 테마코드, 종목명, 종목코드, 설명 등 추출
def get_theme_share_info(theme_code_list):
    import csv
    import time
    import datetime
    from scraping.web_scraping import create_soup

    print("Start get share code")
    print(time.strftime("%Y-%m-%d %H:%M:%S"))
    csv_file = open(f"stock_data\\theme_share_info_{datetime.date.today()}.csv", "w", encoding='utf-8-sig', newline="")
    csv_writer = csv.writer(csv_file)

    # 헤더
    header = (['테마명', '테마코드', '종목명', '종목코드', '종목정보', '종목링크'])
    csv_writer.writerow(header)

    # th_sh_info = [] # 테마명, 테마코드, 테마설명, 종목명, 종목코드, 종목 설명 리스트화
    for th_code in theme_code_list:
        print(f"theme code: {th_code}")
        # time.sleep(3)
        url = f'https://finance.naver.com/sise/sise_group_detail.nhn?type=theme&no={th_code}'
        soup = create_soup(url)
        # 테마명, 테마 설명 가져오기
        th_name = soup.select_one('#contentarea_left > table > tr > td > div > div > strong').get_text()
        th_info = soup.select('#contentarea_left > table > tr > td')[1].p.get_text()
        # th_sh_info.append([th_name, th_code, '', '', th_info]) # 리스트에 추가
        # ws.append([th_name, th_code, '', '', th_info]) # 엑셀에 저장
        csv_writer.writerow([th_name, str(th_code), '', '', th_info])

        # 테마별 종목명, 종목코드
        sh_data = soup.find('tbody').select('tr')[:-2]
        for tag in sh_data:
            # sh_link = "https://finance.naver.com/" + tag.select('td')[0].a['href']
            sh_name = tag.select('td')[0].text
            sh_code = tag.select('td')[0].a['href'].replace("/item/main.nhn?code=", "")
            sh_info = tag.select('td')[1].p.text
            # th_sh_info.append([th_name, th_code, sh_name, sh_code, sh_info, sh_link])  # 리스트에 추가
            # ws.append([th_name, th_code, sh_name, sh_code, sh_info, sh_link]) # 엑셀에 저장
            csv_writer.writerow([th_name, str(th_code), sh_name, str(sh_code), sh_info])

    csv_file.close()
    print("Share code done")
    print(time.strftime("%Y-%m-%d %H:%M:%S"))




'''
3. 실적
'''
# code_dict = {'189980': {'Share Name': '흥국에프엔비', 'Market Code': '031102', 'Market Name': '비알코올음료 및 얼음 제조업'}, '000540': {'Share Name': '흥국화재', 'Market Code': '116501', 'Market Name': '보험업'}, '003280': {'Share Name': '흥아해운', 'Market Code': '085001', 'Market Name': '해상 운송업'}, '037440': {'Share Name': '희림', 'Market Code': '137201', 'Market Name': '건축기술, 엔지니어링 및 관련 기술 서비스업'}, '238490': {'Share Name': '힘스', 'Market Code': '032902', 'Market Name': '특수 목적용 기계 제조업'}}

# 6. 기간별 실적 스크래핑
# share_code_naver.xlsx 파일에서 코드딕트 받아오기
def get_financial_info(code_dict):
    import time
    import datetime
    from bs4 import BeautifulSoup
    from pandas import DataFrame
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    # headless webdriver
    options = webdriver.ChromeOptions()
    options.headless = True
    options.add_argument("window-size=1920x1080")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36")
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.implicitly_wait(3)

    # header = "기간 ,Share Code,Share Name,Market Code,Market Name,매출액,영업이익,영업이익(발표기준),세전계속사업이익,당기순이익,  당기순이익(지배),  당기순이익(비지배),자산총계,부채총계,자본총계,  자본총계(지배),  자본총계(비지배),자본금,영업활동현금흐름,투자활동현금흐름,재무활동현금흐름,CAPEX,FCF,이자발생부채,영업이익률,순이익률,ROE(%),ROA(%),부채비율,자본유보율,EPS(원),PER(배),BPS(원),PBR(배),현금DPS(원),현금배당수익률,현금배당성향(%),발행주식수(보통주)".split(',')

    print("Get Year Financial Info Start : ")
    print(time.strftime("%Y-%m-%d %H:%M:%S"))
    # code_dict = {'016610': {'Share Name': 'DB금융투자 ', 'Market Code': '12', 'Market Name': '증권'}, '005830': {'Share Name': 'DB손해보험 ', 'Market Code': '190', 'Market Name': '손해보험'}, '112190': {'Share Name': 'KC산업 ', 'Market Code': '191', 'Market Name': '상사'}, '000990': {'Share Name': 'DB하이텍 ', 'Market Code': '202', 'Market Name': '반도체와반도체장비'}, '000995': {'Share Name': 'DB하이텍1우 ', 'Market Code': '202', 'Market Name': '반도체와반도체장비'}, '139130': {'Share Name': 'DGB금융지주 ', 'Market Code': '20', 'Market Name': '은행'}, '001530': {'Share Name': 'DI동일 ', 'Market Code': '134', 'Market Name': '섬유,의류,신발,호화품'}, '000210': {'Share Name': 'DL ', 'Market Code': '42', 'Market Name': '건설'}, '000215': {'Share Name': 'DL우 ', 'Market Code': '42', 'Market Name': '건설'}, '375500': {'Share Name': 'DL이앤씨 ', 'Market Code': '42', 'Market Name': '건설'}, '068790': {'Share Name': 'DMS *', 'Market Code': '199', 'Market Name': '디스플레이장비및부품'}, '112190': {'Share Name': 'DRB동일 ', 'Market Code': '36', 'Market Name': '자동차부품'}}

    for key, value in code_dict.items():
        try:
            time.sleep(3)
            print(f"Start: {key}")
            driver.get(f'https://finance.naver.com/item/coinfo.nhn?code={key}')
            wait = WebDriverWait(driver, 3)

            # iframe 접근
            driver.switch_to.frame('coinfo_cp')
            time.sleep(1.5)

            # 연간 분기 돌아가면서 정보 받아오기
            period_dict = {'cns_Tab21': '연간', 'cns_Tab22': '분기'}
            for p_key, p_val in period_dict.items():
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'#{p_key}'))).click()
                time.sleep(1)

                # iframe 에서 html정보 추출
                soup = BeautifulSoup(driver.page_source, "html.parser")

                # 기간 정보
                # period_data = soup.select('table.gHead01')[3].select('thead > tr > th')[2:7]
                # period = [line.text.strip()[:7] for line in period_data]

                period_data = soup.select('table.gHead01')[3].select('thead > tr > th')[2:]
                period = []
                for line in period_data:
                    if line.span.text == "(IFRS연결)":
                        date = line.text.replace("(IFRS연결)", "").strip()
                    elif line.span.text == "(IFRS별도)":
                        date = line.text.replace("(IFRS별도)", "").strip()
                    else:
                        date = line.text.replace("(GAAP개별)", "").strip()
                    period.append(date)
                # print(period_year)

                # financial_dict에 코드정보 입력
                financial_dict = {'Share Code': f'{str(key)}'}
                financial_dict.update(value)  # code_dict의 밸류값 추가

                # 재무정보
                financial_table = soup.select('table.gHead01')[3].select('tbody > tr')
                for line in financial_table:
                    name = line.select_one('th').text  # 재무 항목명
                    data = line.select('td')  # 재무 데이터 : 앞 다섯 기간만, 컨센 제외
                    financial_dict.setdefault(name, [num.text.replace(",", "") for num in data])
                    print(financial_dict)
                # financial_dict를 데이터 프레임으로
                # 컬럼 = list(financial_dict.keys()) = name, 인덱스 = 기간정보
                financial_df = DataFrame(financial_dict, columns=list(financial_dict.keys()), index=period)
                print(financial_df)
                financial_df.to_csv(f"stock_data\\{p_val}실적_raw_{datetime.date.today()}.csv", mode="a", header=False, encoding='utf-8-sig')
        except:
            pass
            print(f'pass: {key}')
    print("Financial Info Done")
    print(time.strftime("%Y-%m-%d %H:%M:%S"))
    driver.quit()



'''
4.etf/etn
'''

# 6. ETF 코드 받기
# wise_compny 사이트에서 etf 코드 가져오기
def get_etf_code():
    from scraping.web_scraping import create_soup

    soup = create_soup("http://comp.wisereport.co.kr/ETF/lookup.aspx")
    table = soup.select_one('table.table').tbody.select('tr')

    etf_code = []
    for tag in table:
        td = tag.select('td')
        if td[0].text == "ETF":
            etf_code.append(td[1].text.strip())
        # else:
        #     etf_list.append([data.text.strip() for data in td])
    print("ETF code done")
    return etf_code

# 7. ETN 코드받기
def get_etn_code():
    from scraping.web_scraping import create_soup

    soup = create_soup("http://comp.wisereport.co.kr/ETF/lookup.aspx")
    table = soup.select_one('table.table').tbody.select('tr')

    etn_code = []
    for tag in table:
        td = tag.select('td')
        if td[0].text == "ETN":
            etn_code.append(td[1].text.strip())
        # else:
        #     etf_list.append([data.text.strip() for data in td])
    print("ETN code done")
    return etn_code

# 7. ETN 코드받기
# def get_etnf_list():
#     from kyu_def.web_scraping import create_soup
#     soup = create_soup("http://comp.wisereport.co.kr/ETF/lookup.aspx")
#     table = soup.select_one('table.table').tbody.select('tr')
#
#     etnf_list = []
#     for tag in table:
#         td = tag.select('td')
#         etnf_list.append([data.text.strip() for data in td])
#
#     return etnf_list


# 8. 네이버 금융에서 개별 etf 정보 가져오기
def get_etf_info(etf_code):
    import time
    import datetime
    import csv
    from scraping.web_scraping import create_soup
    from selenium import webdriver
    from bs4 import BeautifulSoup

    # # CSV
    csv_file = open(f'stock_data\\etf_{datetime.date.today()}.csv', 'w', encoding='utf-8-sig', newline='')
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

    print(f"Get ETF info :")
    print(time.strftime("%Y-%m-%d %H:%M:%S"))

    for code in etf_code:
        time.sleep(3)
        print(f"Start : {code}")
        url = f'https://finance.naver.com/item/coinfo.nhn?code={code}'
        driver.get(url)
        soup = create_soup(url)

        # 이름, 링크, 운용사, 수수료
        etf_name = soup.select_one('#middle > div.h_company > div.wrap_company > h2 > a').text
        etf_link = f'https://finance.naver.com/item/coinfo.nhn?code={code}'
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
            each_etf_info.extend([sh_name, percent])
            # 리스트 진짜 이방법밖에 없는거 실화..? 더 찾아보기
            # each_etf_info.append(sh_name)  # 3차 리스트 저장
            # each_etf_info.append(percent)
        # csv에 입력
        csv_writer.writerow(each_etf_info)

    driver.quit()
    print(f"ETF info completed")
    print(time.strftime("%Y-%m-%d %H:%M:%S"))


# 네이버 업종 밸류옵션값 변경
def get_realtime_value(mk_code_dict):
    import os
    import csv
    import time
    import datetime
    from bs4 import BeautifulSoup
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from stock_scraping_master import get_naver_market_code

    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Color, Side
    import pandas as pd

    '''
    # 1. CSV 파일 생성
    '''
    dir_path = "stock_data/실시간벨류에이션"
    if not os.path.exists(dir_path):
        os.mkdir(dir_path)
    csv_file = open(dir_path + f'\\realtime_stock_value_{datetime.date.today()}.csv', 'w', encoding='utf-8-sig',
                    newline='')
    csv_writer = csv.writer(csv_file)
    csv_writer.writerow(
        ['URL', '종목코드', '종목명', '업종코드', '업종명', '등락률', '시가총액 (억 원)', 'PER', 'ROE', 'PEG', 'ROA', 'PBR', '유보율'])

    '''
    # 2. 네이버 업종 크롤링
    '''
    # headless
    options = webdriver.ChromeOptions()
    options.headless = True
    options.add_argument("window-size=1920x1080")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36")
    driver = webdriver.Chrome(options=options)

    # 마켓코드 받아오기
    # mk_code_dict = get_naver_market_code()

    for mk_name, mk_code in mk_code_dict.items():
        # try:
        # time.sleep(3)
        print(f"Start {mk_name} Category")
        url = f'https://finance.naver.com/sise/sise_group_detail.nhn?type=upjong&no={mk_code}'

        # 셀레니움으로 받아야 옵션 정보가 유지됨
        driver.get(url)
        driver.implicitly_wait(3)
        wait = WebDriverWait(driver, 3)

        titles = driver.find_elements_by_css_selector(
            "#contentarea > div:nth-child(5) > table > thead > tr:nth-child(1) > th")
        title = [line.text for line in titles]
        print(title, "sel")

        # title에 '거래량'이 있을 경우 옵션변경
        if '거래량' in title:

            # 기존 옵션 제거 : 거래량, 매수호가, 거래대금, 매도호가, 전일거래량
            remove_list = [1, 2, 3, 8, 9]
            for num in remove_list:
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'#option{num}'))).click()

            # 원하는 옵션 클릭 : 시가총액, PER, ROE, ROA, PBR, 유보율
            remove_list = [4, 6, 12, 18, 24, 27]
            for num in remove_list:
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'#option{num}'))).click()

            # 옵션 적용 클릭
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.item_btn > a'))).click()  # 적용하기
            time.sleep(1)

        # html 파싱
        soup = BeautifulSoup(driver.page_source, "lxml")
        # 타이틀 다시 확인
        headers = soup.select('#contentarea > div:nth-child(5) > table > thead > tr:nth-child(1) > th')
        head = [line.text for line in headers]
        print(head, "soup")
        table = soup.select('#contentarea > div.box_type_l')[1].select('tbody > tr')[:-2]

        # 종목별 정보 추출 및 기록
        for line in table:
            # 리스트에 이름 및 코드 추가
            share_name = line.td.text
            share_code = line.td.a['href'].replace("/item/main.nhn?code=", "")
            share_link = "https://finance.naver.com/item/main.nhn?code=" + share_code
            info_list = [share_link, str(share_code), share_name, str(mk_code), mk_name]

            # 벨류 데이터 추가 (등락률 부터)
            data = line.select('td')[3:-1]
            for num in data:
                num = num.text.replace(',', '').replace('+', '').replace('%', '')
                if num == '':
                    num = ''
                else:
                    num = float(num)
                info_list.append(num)

            # PEG 밸류 추가
            per = info_list[7]
            roe = info_list[8]
            # per, roe 값을 기준으로 info_list 값 달라짐
            if per == '' or roe == '':  # 값이 없는 경우
                info_list.insert(9, '')
            elif per > 0 and roe > 0:  # 0보다 크거나 같으면 peg 계산
                peg = per / roe
                info_list.insert(9, f"{peg:.1f}")
            elif per <= 0 or roe <= 0:
                info_list.insert(9, '')  # 마이너스인 경우 '' 반환

            # csv에 기록
            csv_writer.writerow(info_list)

        # except:
        #     print(f"err code: {mk_code}, {mk_name}")
        #     pass

    driver.quit()
    csv_file.close()
    print("CSV Completed")

    '''
    # 3. realtime_stock_value.xlsx 파일에 날짜 이름 시트로 붙여넣기
    '''
    # 아예 엑셀파일로 바로 생성하는 것 생각해보기

    file_name = dir_path + "/realtime_stock_value.xlsx"
    df = pd.read_csv(dir_path + f'/realtime_stock_value_{datetime.date.today()}.csv')

    if not os.path.exists(file_name):
        with pd.ExcelWriter(file_name, mode='w', engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=f"{datetime.date.today()}", index=False)
    else:
        with pd.ExcelWriter(file_name, mode='a', engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=f"{datetime.date.today()}", index=False)
    print("df to 엑셀")


    '''
    # 4. 엑셀 서식 추가
    '''
    # 셀서식
    align_center = Alignment(horizontal="center", vertical="center")
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"),
                         bottom=Side(style="thin"))

    # 글꼴 서식
    font_bold = Font(name="맑은 고딕", size=11, bold=True)
    font_red = Font(name="맑은 고딕", size=11, color="FE2E2E")
    font_blue = Font(name="맑은 고딕", size=11, color="1E88E5")

    # color_fill
    orange_fill = PatternFill(patternType="solid", fgColor=Color("FEF5E7"))
    dark_pink_fill = PatternFill(patternType="solid", fgColor=Color("EC7063"))
    dark_blue_fill = PatternFill(patternType="solid", fgColor=Color("5DADE2"))

    # 워크북 불러오기
    wb = load_workbook(file_name)
    ws = wb[f"{datetime.date.today()}"]

    # 컬럼너비 (종목명, 업종명, 상승종목 수, 하락종목 수)
    width_dict = {"C": 21, "E": 19, "O": 13, "P": 13, "Q": 13}
    for key, value in width_dict.items():
        ws.column_dimensions[f"{key}"].width = value

    # 헤더 서식
    for row in ws["A1:M1"]:
        for cell in row:
            cell.font = font_bold
            cell.alignment = align_center
            cell.fill = orange_fill
            cell.border = thin_border

    # 등락률 글자색
    count_inc = 0  # 상승 종목 갯수
    count_dec = 0  # 하락 종목 갯수
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

    # 상승 하락 종목 셀 서식
    ws["O1"].value = "상승 종목 수"
    ws["P1"].value = "하락 종목 수"
    ws["Q1"].value = "총 종목 수"
    ws["O1"].fill = dark_pink_fill
    ws["P1"].fill = dark_blue_fill
    ws["Q1"].fill = orange_fill
    ws["O2"].value = count_inc
    ws["P2"].value = count_dec
    ws["Q2"].value = count_inc + count_dec

    for row in ws[f"O1:Q1"]:
        for cell in row:
            cell.font = font_bold

    for row in ws[f"O1:Q2"]:
        for cell in row:
            cell.border = thin_border
            cell.alignment = align_center

    print(count_inc, count_dec, count_inc + count_dec)

    # 전체 보더
    for row in ws[f"A2:M{ws.max_row}"]:
        for cell in row:
            cell.border = thin_border

    # 오토필터
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = 'A2'

    # 컬럼 숨기기 (URL)
    ws.column_dimensions["A"].hidden = True

    wb.save(file_name)
    print("엑셀 서식 완료")


def get_market_fluctuation():
    import csv
    import datetime
    from scraping.web_scraping import create_soup

    csv_file = open(f"stock_data\\업종데일리등락률\\market_fluctuation_{datetime.date.today()}.csv", "a", encoding='utf-8-sig', newline="")
    csv_writer = csv.writer(csv_file)

    # period = datetime.date.today()

    # 헤더
    header = (['URL', '마켓명', '마켓코드', '전일대비 등락률'])
    csv_writer.writerow(header)

    # 종목코드 종목명 전일대비등락률
    # theme_fluc_list = []
    print("Start get market code")


    # 네이버 금융 업종
    url = "https://finance.naver.com/sise/sise_group.nhn?type=upjong"
    soup = create_soup(url)
    theme_data = soup.select_one("table.type_1").select("tr")[3:]
    for line in theme_data:
        td = line.select('td')
        if len(td) <= 1:  # 의미없는 자료 제거
            continue
        market_name = td[0].text.strip()
        market_code = td[0].a["href"].replace("/sise/sise_group_detail.nhn?type=upjong&no=", "")
        fluc = td[1].text.strip().replace('%', '').replace('+', '')
        link = "https://finance.naver.com/sise/sise_group_detail.nhn?type=theme&no=" + market_code

        csv_writer.writerow([link, market_name, market_code, fluc])

    print("csv done")


def get_theme_fluctuation():
    import csv
    import datetime
    from scraping.web_scraping import create_soup

    csv_file = open(f"stock_data\\테마데일리등락률\\theme_fluctuation_{datetime.date.today()}.csv", "a", encoding='utf-8-sig', newline="")
    csv_writer = csv.writer(csv_file)

    # period = datetime.date.today()

    # 헤더
    header = (['URL', '테마명', '테마코드', '전일대비 등락률'])
    csv_writer.writerow(header)

    # 테마명 테마코드 전일대비등락률
    # theme_fluc_list = []
    print("Start get theme code")
    for page in range(1, 7):
        print(f"On page : {page}")

        # 네이버 금융 테마별 시세
        url = f"https://finance.naver.com/sise/theme.nhn?&page={page}"
        soup = create_soup(url)
        theme_data = soup.select_one("table.type_1").select("tr")[3:]
        for line in theme_data:
            td = line.select('td')
            if len(td) <= 1:  # 의미없는 자료 제거
                continue
            theme_name = td[0].text.strip()
            theme_code = td[0].a["href"].replace("/sise/sise_group_detail.nhn?type=theme&no=", "")
            fluc = td[1].text.strip().replace('%', '').replace('+', '')
            link = "https://finance.naver.com/sise/sise_group_detail.nhn?type=upjong&no=" + theme_code

            csv_writer.writerow([link, theme_name, theme_code, fluc])

    print("csv done")

'''# 8. 네이버 금융에서 개별 etf 정보 가져오기
def get_etf_info(etf_code):
    import time
    import csv
    from kyu_def.web_scraping import create_soup, create_csv
    from selenium import webdriver
    from bs4 import BeautifulSoup
    from openpyxl import Workbook

    # csv 테스트
    title = 'etf_etn'
    header = ['ETF이름', '링크', 'ETF코드', '시가총액(억 원)', '운용사', '수수료', '수익률', '', '', '', '구성종목']
    csv_file = create_csv(title, header)
    csv_writer = csv.writer(csv_file)
    csv_writer.writerow(['', '', '', '', '', '', '1개월', '3개월', '6개월', '1년'])

    # # 워크북 생성
    # wb = Workbook()
    # ws = wb.active
    # ws.title = 'ETF'

    # # 엑셀 헤더
    # ws.append(['ETF이름', '링크', 'ETF코드', '시가총액(억 원)', '운용사', '수수료', '수익률', '', '', '', '구성종목'])
    # ws.append(['', '', '', '', '', '', '1개월', '3개월', '6개월', '1년'])

    # 헤드리스 셀레니움
    options = webdriver.ChromeOptions()
    options.headless = True
    options.add_argument("window-size=1920x1080")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36")
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.implicitly_wait(3)

    # etf 정보를 리스트화
    etf_info = []
    for code in etf_code:
        time.sleep(3)
        url = f'https://finance.naver.com/item/coinfo.nhn?code={code}'
        driver.get(url)
        soup = create_soup(url)

        # 이름, 링크, 운용사, 수수료
        etf_name = soup.select_one('#middle > div.h_company > div.wrap_company > h2 > a').text
        etf_link = f'https://finance.naver.com/item/coinfo.nhn?code={code}'
        company = soup.select_one('#tab_con1').select_one('div:nth-child(4)').select('td')[1].text
        commission = soup.select_one('#tab_con1').select_one('div:nth-child(4)').select('td')[0].em.text.replace('%', '')
        each_etf_info = [etf_name, etf_link, code, company, commission] # 1차 리스트 저장

        # iframe으로 이동
        driver.switch_to.frame('coinfo_cp')
        time.sleep(0.5)
        soup_iframe = BeautifulSoup(driver.page_source, "html.parser")

        # 시가총액
        capital = soup_iframe.select_one('#status_grid_body').select('tr')[5].td.text.replace(',', '')[:-2]
        each_etf_info.append(capital)

        # 수익률
        earnings = soup_iframe.select_one('#status_grid_body').select('tr')[-1].td.select('span')
        for line in earnings:
            earning = line.text.replace('%', '').replace('+', '')
            if earning == "N/A":
                earning = ''
            elif earning == "":
                earning = ''
            else:
                earning = earning
            each_etf_info.append(earning)

        # 구성종목
        top_10 = soup_iframe.select_one('#CU_grid_body').select('tr')[:10]
        for line in top_10:
            sh_name = line.select('td')[0].text
            percent = line.select('td')[2].text
            if percent == '-':
                percent = ''
            else:
                percent = float(percent)
            # 리스트 진짜 이방법밖에 없는거 실화..? 더 찾아보기
            each_etf_info.append(sh_name)  # 3차 리스트 저장
            each_etf_info.append(percent)

            # iframe 나오기
            driver.switch_to.default_content()

        etf_info.append(each_etf_info)  # 각각의 etf 정보를 토탈 리스트에 넣어줌
        ws.append(each_etf_info)  # 엑셀에 입력

    driver.quit()
    return etf_info'''