import requests
from bs4 import BeautifulSoup
import os
import datetime

keyword = input("검색어 : ")
url = f"https://news.joins.com/Search/TotalNews?page=30&Keyword={keyword}&SortType=New&SearchCategoryType=TotalNews"
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.72 Safari/537.36 Edg/89.0.774.45"}
response = requests.get(url, headers=headers)
soup = BeautifulSoup(response.text, "html.parser")

table = soup.select(".section_news h2.mg a")

for ele in table:
    # if len(table) == 0:
    #     break
    title = ele.text
    link = ele['href']
    news_req = requests.get(link).text
    news_soup = BeautifulSoup(news_req, "html.parser")
    content = news_soup.select_one("div#article_body").text.strip().replace("     ", "")
    print(f"""
    제목 : {title}
    링크 : {link}
    {content}
    """)
    print()