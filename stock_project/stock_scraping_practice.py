# This Python file uses the following encoding: utf-8
import os, sys

from scraping.stock_scraping_master import get_code_from_excel

# def concat_financial_info():
#
#     import datetime
#     import pandas as pd
#
#     pd.options.display.max_columns = 50
#     pd.options.display.max_rows = 500
#
#     header = ['기간', '종목코드', '종목명', '업종코드', '업종명', '매출액', 'salesVar', '영업이익', 'opVar', '영업이익(발표기준)', '세전계속사업이익', '당기순이익', 'npVar', '당기순이익(지배)', '당기순이익(비지배)',
#               '자산총계', '부채총계', '자본총계', '자본총계(지배)', '자본총계(비지배)', '자본금', '영업활동현금흐름', '투자활동현금흐름', '재무활동현금흐름', 'CAPEX', 'FCF', '이자발생부채', '영업이익률', '순이익률',
#               'ROE(%)', 'ROA(%)', '부채비율', '자본유보율', 'EPS(원)', 'PER(배)', 'BPS(원)', 'PBR(배)', '현금DPS(원)', '현금배당수익률', '현금배당성향(%)', '발행주식수(보통주)']
#
#     # writer = pd.ExcelWriter(f"stock_data/재무실적/종목실적_raw_{datetime.date.today()}.xlsx")
#     # 기간별 재무 정보 불러오기
#     for period in ['연간', '분기']:
#         print(f'{period} 시작')
#
#         df = pd.read_csv(f"stock_data/재무실적/로우데이터/{period}실적_raw_2021-03-31.csv", names=header)  # new 재무자료
#         old_df = pd.read_excel(f"stock_data/재무실적/로우데이터/{period}실적_통합_2021-04-01.xlsx")  # pre 재무자료 , dtype={'종목코드': str}
#
#         concat_df = pd.concat([df, old_df])  # 이전 재무 데이터와 합치기
#         concat_df.dropna(subset=['당기순이익'], inplace=True)  # 당기순이익이 비어있는 행 제거
#         concat_df.drop_duplicates(['기간', '종목코드'], inplace=True)  # 중복 제거
#         concat_df.sort_values(by=['종목명', '기간'], inplace=True)  # 종목명과 기간으로 정렬
#
#         # 엑셀로 저장
#         concat_df.to_excel(f"stock_data/재무실적/로우데이터/{period}실적_통합_{datetime.date.today()}.xlsx")
#         print(f'{period} 실적 저장 완료')
# concat_financial_info()
    # writer.save()

# def concat_financial_info(old_file_path): # 재무실적_통합
#
#     import datetime
#     import pandas as pd
#
#     pd.options.display.max_columns = 50
#     pd.options.display.max_rows = 500
#
#     header = ['기간', '종목코드', '종목명', '업종코드', '업종명', '매출액', 'salesVar', '영업이익', 'opVar', '영업이익(발표기준)', '세전계속사업이익', '당기순이익', 'npVar', '당기순이익(지배)', '당기순이익(비지배)',
#               '자산총계', '부채총계', '자본총계', '자본총계(지배)', '자본총계(비지배)', '자본금', '영업활동현금흐름', '투자활동현금흐름', '재무활동현금흐름', 'CAPEX', 'FCF', '이자발생부채', '영업이익률', '순이익률',
#               'ROE(%)', 'ROA(%)', '부채비율', '자본유보율', 'EPS(원)', 'PER(배)', 'BPS(원)', 'PBR(배)', '현금DPS(원)', '현금배당수익률', '현금배당성향(%)', '발행주식수(보통주)']
#
#     # writer = pd.ExcelWriter(f"stock_data/재무실적/종목실적_raw_{datetime.date.today()}.xlsx")
#     for period in ['연간', '분기']:
#
#         # 기간별 재무 정보 불러오기
#         print(f'{period} 시작')
#         df = pd.read_csv(f"stock_data/재무실적/로우데이터/{period}실적_raw_2021-03-31.csv", names=header)  # new 재무자료
#         print(df.head(10))
#         old_df = pd.read_excel(old_file_path, dtype={'종목코드': str}, sheet_name=f'{period}', index_col=None)  # pre 재무자료
#         print(old_df.head(10))
#         # print(df.head(50))
#
#         concat_df = pd.concat([df, old_df])  # 이전 재무 데이터와 합치기
#         concat_df.dropna(subset=['당기순이익'], inplace=True)  # 당기순이익이 비어있는 행 제거
#         concat_df.drop_duplicates(['기간', '종목코드'], inplace=True)  # 중복 제거
#         concat_df.sort_values(by=['종목명', '기간'], inplace=True)  # 종목명과 기간으로 정렬
#         print(concat_df.head(50))
#
#         # 엑셀로 저장
#         concat_df.to_excel(f"stock_data/재무실적/로우데이터/재무실적_통합_{datetime.date.today()}.xlsx", index=None)
#         print(f'{period} 실적 엑셀 저장 완료')
#
# concat_financial_info('stock_data/재무실적/로우데이터/재무실적_통합_2021-04-01.xlsx')

# def clean_quarter_period(a):
#     if 'E' in a:
#         pass
#     elif '/11' in a or '/10' in a:
#         a = a[:5] + '12'
#     elif '/08' in a or '/07' in a:
#         a = a[:5] + '09'
#     elif '/05' in a or '/04' in a:
#         a = a[:5] + '06'
#     elif '/02' in a or '/01' in a:
#         a = a[:5] + '03'
#     return a
#
# def create_total_data_excel():
#
#     import pandas as pd
#     import datetime
#     from openpyxl import load_workbook
#     from openpyxl.styles import Font, Alignment, PatternFill, Border, Color, Side
#     pd.options.display.max_columns = 50
#     pd.options.display.max_rows = 500
#
#     ################
#     # 1. 파일합치기
#     ################
#     writer = pd.ExcelWriter(f"stock_data/통합데이터/주식정보_통합_{datetime.date.today()}.xlsx")
#     # 1-1 재무데이터
#     for period in ['연간', '분기']:
#         df = pd.read_excel(f'stock_data/재무실적/로우데이터/{period}실적_통합_2021-04-01.xlsx')
#         print(f"{period} 실적 저장")
#
#         # 기간 인덱스 정리
#         for i in range(len(df['기간'])):
#             if period == '연간': # 2019, 2020, 2021
#                 df.iloc[i, 0] = list(map(lambda a: a[:4] + '(E)' if 'E' in a else a[:4], [df.iloc[i, 0]]))
#             else: # /03, /06, /09, /12
#                 df.iloc[i, 0] = clean_quarter_period(df.iloc[i, 0])
#         df.to_excel(writer, sheet_name=f'{period}실적', index=None)
#
#     # 1-2 등락률
#     for item in ['업종', '테마']:
#         fluc_df = pd.read_excel("stock_data/데일리등락률_통합.xlsx", sheet_name=f'{item}')
#         fluc_df.to_excel(writer, sheet_name=f'{item}등락률', index=None)
#         print(f"{item} 등락률 저장")
#
#     # 1-3 실시간밸류
#     value_df = pd.read_excel("stock_data/실시간벨류에이션/realtime_stock_value.xlsx")  # csv로 변경
#     value_df.to_excel(writer, sheet_name="실시간밸류", index=None)
#     print("실시간밸류 저장")
#     writer.save()
# #
#     ###############
#     # 2. 엑셀다듬기
#     ###############
#     # 셀서식
#     align_center = Alignment(horizontal="center", vertical="center")
#     thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"),
#                          bottom=Side(style="thin"))
#
#     # 글꼴 서식
#     font_normal = Font(name="맑은 고딕", size=11)
#     font_bold = Font(name="맑은 고딕", size=11, bold=True)
#     font_red = Font(name="맑은 고딕", size=11, color="FE2E2E")
#     font_blue = Font(name="맑은 고딕", size=11, color="1E88E5")
#
#     # color_fill
#     orange_fill = PatternFill(patternType="solid", fgColor=Color("FEF5E7"))
#     green_fill = PatternFill(patternType="solid", fgColor=Color("BBE8A2"))
#     dark_orange_fill = PatternFill(patternType="solid", fgColor=Color("FC9C30"))
#
#     light_pink_fill = PatternFill(patternType="solid", fgColor=Color("ffe6e6"))
#     pink_fill = PatternFill(patternType="solid", fgColor=Color("ffcccc"))
#     pink_fill_2 = PatternFill(patternType="solid", fgColor=Color("ffb3b3"))
#     pink_fill_3 = PatternFill(patternType="solid", fgColor=Color("ff8080"))
#     dark_pink_fill = PatternFill(patternType="solid", fgColor=Color("ff6666"))
#     red_fill = PatternFill(patternType="solid", fgColor=Color("ff0000"))
#
#     light_blue_fill = PatternFill(patternType="solid", fgColor=Color("e6f5ff"))
#     blue_fill = PatternFill(patternType="solid", fgColor=Color("ccebff"))
#     blue_fill_2 = PatternFill(patternType="solid", fgColor=Color("99d6ff"))
#     blue_fill_3 = PatternFill(patternType="solid", fgColor=Color("66c2ff"))
#     dark_blue_fill = PatternFill(patternType="solid", fgColor=Color("4db8ff"))
#     navy_fill = PatternFill(patternType="solid", fgColor=Color("008ae6"))
#
#     # 2-1 통합 데이터 엑셀 불러오기
#     wb = load_workbook(f"stock_data/통합데이터/주식정보_통합_{datetime.date.today()}.xlsx")
#
#     # 2-2 재무
#     col_dict = {'연간':green_fill, '분기':dark_orange_fill}
#     for item2 in col_dict:
#
#         ws = wb[f'{item2}실적']
#
#         # 헤더 서식
#         for y in range(1, ws.max_column + 1):
#             ws.cell(row=1, column=y).font = font_bold
#             ws.cell(row=1, column=y).alignment = align_center
#             ws.cell(row=1, column=y).fill = col_dict[item2]
#
#         # 오토필터
#         ws.auto_filter.ref = ws.dimensions
#         ws.freeze_panes = 'D2'
#         print(f"{item2} 엑셀 서식 완료")
#
#
#     # 2-3 등락률
#     for item3 in ["업종", "테마"]:
#         ws = wb[f'{item3}등락률']
#
#         # 컬럼너비
#         ws.column_dimensions["B"].width = 30
#
#         # 헤더 서식
#         for y in range(1, ws.max_column + 1):
#             ws.cell(row=1, column=y).font = font_bold
#             ws.cell(row=1, column=y).alignment = align_center
#             ws.cell(row=1, column=y).fill = orange_fill
#
#         # 색칠
#         for x in range(2, ws.max_row+1):
#             for y in range(3, ws.max_column+1):
#                 if ws.cell(row=x, column=y).value == None:
#                     continue
#                 if ws.cell(row=x, column=y).value > 5:
#                     ws.cell(row=x, column=y).fill = red_fill
#                 elif ws.cell(row=x, column=y).value > 4:
#                     ws.cell(row=x, column=y).fill = dark_pink_fill
#                 elif ws.cell(row=x, column=y).value > 3:
#                     ws.cell(row=x, column=y).fill = pink_fill_3
#                 elif ws.cell(row=x, column=y).value > 2:
#                     ws.cell(row=x, column=y).fill = pink_fill_2
#                 elif ws.cell(row=x, column=y).value > 1:
#                     ws.cell(row=x, column=y).fill = pink_fill
#                 elif ws.cell(row=x, column=y).value > 0:
#                     ws.cell(row=x, column=y).fill = light_pink_fill
#                 elif ws.cell(row=x, column=y).value > -1:
#                     ws.cell(row=x, column=y).fill = light_blue_fill
#                 elif ws.cell(row=x, column=y).value > -2:
#                     ws.cell(row=x, column=y).fill = blue_fill
#                 elif ws.cell(row=x, column=y).value > -3:
#                     ws.cell(row=x, column=y).fill = blue_fill_2
#                 elif ws.cell(row=x, column=y).value > -4:
#                     ws.cell(row=x, column=y).fill = dark_blue_fill
#                 else:
#                     ws.cell(row=x, column=y).fill = navy_fill
#
#         # 전체 보더
#         for x in range(1, ws.max_row + 1):
#             for y in range(1, ws.max_column + 1):
#                 ws.cell(row=x, column=y).border = thin_border
#
#         # 오토필터
#         ws.auto_filter.ref = ws.dimensions
#         ws.freeze_panes = 'C2'
#
#         # 하이퍼링크
#         hyper_dict = {"업종": "https://finance.naver.com/sise/sise_group_detail.nhn?type=upjong&no=",
#                       "테마": "https://finance.naver.com/sise/sise_group_detail.nhn?type=theme&no="}
#         for i in range(2, ws.max_row):
#             ws[f"B{i}"].hyperlink = hyper_dict[f"{item3}"] + str(ws[f"A{i}"].value)
#
#         print(f"{item3} 엑셀 서식 완료")
#
#     # 2-4 실시간 밸류
#     ws = wb['실시간밸류']
#
#     # 컬럼너비 (종목명, 업종명, 상승종목 수, 하락종목 수)
#     width_dict = {"C": 21, "E": 19, "O": 13, "P": 13, "Q": 13}
#     for key, value in width_dict.items():
#         ws.column_dimensions[f"{key}"].width = value
#
#     # 헤더 서식
#     for row in ws["A1:M1"]:
#         for cell in row:
#             cell.font = font_bold
#             cell.alignment = align_center
#             cell.fill = orange_fill
#             cell.border = thin_border
#
#     # 등락률 글자색
#     count_inc = 0  # 상승 종목 갯수
#     count_dec = 0  # 하락 종목 갯수
#     for col in ws[f"F2:F{ws.max_row}"]:
#         for cell in col:
#             if cell.value > 0:
#                 cell.font = font_red
#                 count_inc += 1
#             else:
#                 cell.font = font_blue
#                 count_dec += 1
#
#     print("상승종목 수 :", count_inc, "하락종목 수 :", count_dec, "총 종목 수 :", count_inc + count_dec)
#
#     # # 상승 하락 종목 셀 서식
#     # ws["O1"].value = "상승 종목 수"
#     # ws["P1"].value = "하락 종목 수"
#     # ws["Q1"].value = "총 종목 수"
#     # ws["O1"].fill = dark_pink_fill
#     # ws["P1"].fill = dark_blue_fill
#     # ws["Q1"].fill = orange_fill
#     # ws["O2"].value = count_inc
#     # ws["P2"].value = count_dec
#     # ws["Q2"].value = count_inc + count_dec
#     # for row in ws[f"O1:Q2"]:
#     #     for cell in row:
#     #         cell.border = thin_border
#     #         cell.alignment = align_center
#     # for row in ws[f"O1:Q1"]:
#     #     for cell in row:
#     #         cell.font = font_bold
#
#     # 하이퍼링크
#     for i in range(2, ws.max_row):
#         ws[f"C{i}"].hyperlink = ws[f"A{i}"].value
#
#
#     # 전체 보더
#     for row in ws[f"A2:M{ws.max_row}"]:
#         for cell in row:
#             cell.border = thin_border
#
#     # 오토필터
#     ws.auto_filter.ref = ws.dimensions
#     ws.freeze_panes = 'A2'
#
#     # 컬럼 숨기기 (URL)
#     ws.column_dimensions["A"].hidden = True
#     print("밸류 엑셀 서식 완료")
#
#     wb.save(f"stock_data/통합데이터/주식정보_통합_{datetime.date.today()}.xlsx")

# create_total_data_excel()



# {'060310':{'종목명':'3S', '업종코드':314, '업종명':'디스플레이', }}
import pandas as pd
# year = pd.MultiIndex.from_product([col_list, ['연간'], ['2016', '2017', '2018', '2019', '2020', '2015', '2021(E)', '2022(E)', '2023(E)', '2020(E)', '2014']])
# quarter = pd.MultiIndex.from_product([col_list, ['분기'], ['2019/09', '2019/12', '2020/03', '2020/06', '2020/09', '2020/12', '2021/03(E)', '2021/06(E)', '2021/09(E)', '2020/12(E)']])
#

df_list = []
col_list = ['기간', '종목코드', '종목명', '업종코드', '업종명', '매출액', 'salesVar', '영업이익', 'opVar', '당기순이익', 'npVar', '영업활동현금흐름', '투자활동현금흐름', '재무활동현금흐름']
for item in ['연간', '분기']:
    print(f"{item}")
    df = pd.read_excel('stock_data/통합데이터/주식정보_통합_2021-04-03.xlsx', sheet_name=f'{item}실적', index_col=None)
    # index = pd.MultiIndex.from_product([col_list, [f'{item}'], df['기간'].unique()])
    df_list.append(df[col_list])

concat_df = pd.concat(df_list)
concat_df.to_excel("test.xlsx", index=None)
print(concat_df)





    # print(df.loc['060310', [col_list]])







        # print(df.iloc[i, 0])
    # sorted_id = sorted(id_list)
    # print(sorted_id)
    #
    # cleaned_id = list(map(lambda a: a[:4] + '(E)' if re.compile('(E)').search(a) else a[:4], sorted_id))

    # for num, index in enumerate(id_list):
    #     print(num, index)
    #     for col in col_dict:
    #         data = df.loc[f'{index}', [f'{col}']]
    #         if len(data) == 1: # 1줄짜리 데이터 (시리즈) 패스
    #             break
    #         data_list = [data.iloc[i, 0] for i in range(len(data))]
    #         i_list = list(range(1, len(data)))
    #         gr_list = get_growth_rate(data_list, i_list)
    #         gr_list.insert(0, '')
    #         df.loc[f'{index}', [col_dict[col]]] = gr_list
    #
    # df.to_excel(f'stock_data/재무실적/로우데이터/{period}실적_통합_2021-04-01_test3.xlsx')
    # print(f"{period} 실적 완료")





        # print(df[col_dict[col]])
#         df = df[
#             ['Share Code', 'Share Name', 'Market Code', 'Market Name', '매출액', 'salesVar', '영업이익', 'opVar',
#              '영업이익(발표기준)', '세전계속사업이익', '당기순이익', 'npVar', '당기순이익(지배)', '당기순이익(비지배)', '자산총계', '부채총계', '자본총계',
#              '자본총계(지배)', '자본총계(비지배)', '자본금', '영업활동현금흐름', '투자활동현금흐름', '재무활동현금흐름', 'CAPEX', 'FCF', '이자발생부채', '영업이익률',
#              '순이익률', 'ROE(%)', 'ROA(%)', '부채비율', '자본유보율', 'EPS(원)', 'PER(배)', 'BPS(원)', 'PBR(배)', '현금DPS(원)',
#              '현금배당수익률', '현금배당성향(%)', '발행주식수(보통주)']]





    # id_list = df.index.tolist()
    # for index in id_list:
    #     for col in {'매출액':'salesVar', '영업이익':'opVar', '당기순이익':'npVar'}:
    #         data = df.loc[f'{index}', [f'{col}']]
    #         data = data.to_numpy()
    #         print(data)
    #         print(type(data))
    #         print(df.loc['265520', [f'{col}']])
    #         print(type(df.loc['265520', [f'{col}']]))


# print(financial_df)
# 컬럼 순서 조정
# financial_df = financial_df[
#     ['Share Code', 'Share Name', 'Market Code', 'Market Name', '매출액', 'salesVar', '영업이익', 'opVar',
#      '영업이익(발표기준)', '세전계속사업이익', '당기순이익', 'npVar', '당기순이익(지배)', '당기순이익(비지배)', '자산총계', '부채총계', '자본총계',
#      '자본총계(지배)', '자본총계(비지배)', '자본금', '영업활동현금흐름', '투자활동현금흐름', '재무활동현금흐름', 'CAPEX', 'FCF', '이자발생부채', '영업이익률',
#      '순이익률', 'ROE(%)', 'ROA(%)', '부채비율', '자본유보율', 'EPS(원)', 'PER(배)', 'BPS(원)', 'PBR(배)', '현금DPS(원)',
#      '현금배당수익률', '현금배당성향(%)', '발행주식수(보통주)']]
# print(financial_df)
