import requests
from bs4 import BeautifulSoup as bs
from BlbieDictionary import bible

class BibleCrawler:
    def get_bible(self, book, chap, start_verse, end_verse):
        bible_list = []
        page = requests.get(
            "https://www.bskorea.or.kr/bible/korbibReadpage.php?version=GAE&book={}&chap={}&sec={}&cVersion"
            "=&fontSize=15px&fontWeight=normal#focus".format(bible[book], chap, start_verse))
        soup = bs(page.text, "html.parser")

        elements = soup.select('#tdBible1 > span')

        # elements의 각 요소는 [숫자][공백3개][내용]의 형태
        for index in range(start_verse, end_verse + 1):
            transformed_string = elements[index].text.replace(u'\xa0', u' ')  # nbsp 제거하기  
            transformed_string = transformed_string[2:len(transformed_string)] # 앞에서부터 3자리 지우기
            transformed_string = transformed_string.strip() # 앞뒤 공백제거
            transformed_string = " " + transformed_string # 맨앞 공백1개 추가
            bible_list.append(transformed_string)

        return bible_list
