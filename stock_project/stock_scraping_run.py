from scraping.stock_scraping_master import get_etf_code, get_etf_info

# themecode 잘 작동
# theme_code_list = get_naver_theme_code()
# get_theme_share_info(theme_code_list)

# etf 잘 작동
etf_code = get_etf_code()
get_etf_info(etf_code)
# etn 코드 받기 잘 작동
# etn_code = get_etn_code()
# etn_info 받기 코드 다시 짜야함 페이지 구성이 다름
# print(etn_code)

