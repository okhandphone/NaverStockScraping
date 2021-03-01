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
    wb = load_workbook("C:\\Users\\MING\\PycharmProjects\\web_automation\\stock_project\\krxcode.xlsx")  # 불러올 파일명 넣기
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
    wb.save("mk_sh_code_naver.xlsx")
    print("Total Code file Completed")

# 3. 엑셀에서 업종코드 딕셔너리로 가져오기
def get_code_from_excel():
    from openpyxl import load_workbook  # 파일 불러오기
    # 상대경로로 수정해야함
    wb = load_workbook("C:\\Users\\MING\\PycharmProjects\\web_automation\\stock_project\\mk_sh_code_naver.xlsx")  # 불러올 파일명 넣기
    ws = wb.active

    code_dict = {}
    for x in range(2, ws.max_row + 1):
        code = [ws.cell(row=x, column=y).value for y in range(1, ws.max_column + 1)]
        code_dict.setdefault(f"{code[0]}", {"Share Name": code[1], "Market Code": str(code[2]), "Market Name": code[3]})
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
    csv_file = open(f"theme_share_info_{datetime.date.today()}.csv", "w", encoding='utf-8-sig', newline="")
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

    # header = ",Share Code,Share Name,Market Code,Market Name,매출액,영업이익,영업이익(발표기준),세전계속사업이익,당기순이익,  당기순이익(지배),  당기순이익(비지배),자산총계,부채총계,자본총계,  자본총계(지배),  자본총계(비지배),자본금,영업활동현금흐름,투자활동현금흐름,재무활동현금흐름,CAPEX,FCF,이자발생부채,영업이익률,순이익률,ROE(%),ROA(%),부채비율,자본유보율,EPS(원),PER(배),BPS(원),PBR(배),현금DPS(원),현금배당수익률,현금배당성향(%),발행주식수(보통주)".split(',')

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
                period_data = soup.select('table.gHead01')[3].select('thead > tr > th')[2:7]
                period = [line.text.strip()[:7] for line in period_data]

                # financial_dict에 코드정보 입력
                financial_dict = {'Share Code': f'{str(key)}'}
                financial_dict.update(value)  # code_dict의 정보를 추가

                # 재무정보
                financial_table = soup.select('table.gHead01')[3].select('tbody > tr')
                for line in financial_table:
                    name = line.select_one('th').text  # 재무 항목명
                    data = line.select('td')[:5]  # 재무 데이터 : 앞 다섯 기간만, 컨센 제외
                    financial_dict.setdefault(name, [num.text.replace(",", "") for num in data])

                # financial_dict를 데이터 프레임으로
                # 컬럼 = list(financial_dict.keys()) = name, 인덱스 = 기간정보
                financial_df = DataFrame(financial_dict, columns=list(financial_dict.keys()), index=period)
                financial_df.to_csv(f"{p_val}실적_raw_{datetime.date.today()}.csv", mode="a", header=False, encoding='utf-8-sig')
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
    csv_file = open(f'etf_{datetime.date.today()}.csv', 'w', encoding='utf-8-sig', newline='')
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


# etf_code = get_etf_code()
# get_etf_info(etf_code)


# 네이버 업종 페이지에서 실시간 밸류팩터들 추출
# 번갈아 가면서 작동 ㄴㄴ
def get_real_time_value(mk_code_dict):
    import time
    import csv
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from bs4 import BeautifulSoup


    # CSV
    csv_file = open('종목 실시간 밸류에이션.csv', 'a', encoding='utf-8-sig', newline='')
    csv_writer = csv.writer(csv_file)
    csv_writer.writerow(['종목코드', '종목명', '업종코드', '업종명', '시가총액', '외인비율', 'PER', 'ROE', 'ROA', 'PBR'])

    #  headless
    options = webdriver.ChromeOptions()
    options.headless = True
    options.add_argument("window-size=1920x1080")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36")
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()

    #  mk_code_dict를 파라미터로
    for key, value in mk_code_dict.items():
        try:
            time.sleep(5)
            print(f"Start {value} Category")
            driver.get(f'https://finance.naver.com/sise/sise_group_detail.nhn?type=upjong&no={value}')
            wait = WebDriverWait(driver, 3)

            # 기존 옵션 제거
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option1'))).click()  #  거래량
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option2'))).click()  #  매수호가
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option3'))).click()  # 거래대금
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option8'))).click()  #  매도호가
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option9'))).click()  #  전일 거래량

            # 원하는 옵션 클릭
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option4'))).click()  #  시가총액
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option6'))).click()  #  PER
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option12'))).click()  #  ROE
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option18'))).click()  #  ROA
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option24'))).click()  #  PBR
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#option27'))).click()  # 유보율

            # 옵션 적용 클릭
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.item_btn > a'))).click()  # 적용하기

            time.sleep(1.5)
            # driver.save_screenshot(f"screenshot{key}.png")
            # 실시간 기업 밸류에이션 스크래핑
            soup = BeautifulSoup(driver.page_source, "html.parser")
            table = soup.select('#contentarea > div.box_type_l')[1].select('tbody > tr')[:-2]
            for line in table:
                share_name = line.td.text
                share_code = line.td.a['href'].replace("/item/main.nhn?code=", "")
                info_list = [share_code, share_name, value, key]

                data = line.select('td')[4:-1]
                for num in data:
                    num = num.text.replace(',', '')
                    info_list.append(num)
                csv_writer.writerow(info_list)
              # ws.append(info_list)

        except:
            print(f"pass: {value}")
            pass
    # wb.save("종목 실시간 벨류에이션.xlsx")
    driver.quit()
    csv_file.close()
    print("Excel file Completed")


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