import requests
from bs4 import BeautifulSoup as bs


class BibleCrawler:

    # def __init__(self, book, chap, sec, endSec):
    #     self.book = book
    #     self.chap = chap
    #     self.sec = sec
    #     self.endSec = endSec

    def getBible(self, book, chap, sec, endSec):
        bibleList = []
        page = requests.get(
            "https://www.bskorea.or.kr/bible/korbibReadpage.php?version=GAE&book={}&chap={}&sec={}&cVersion"
            "=&fontSize=15px&fontWeight=normal#focus".format(book, chap, sec))
        soup = bs(page.text, "html.parser")

        elements = soup.select('#tdBible1 > span')

        # 23   하나님이 우리를 사랑하사~~ 처럼 "숫자몇자리 + 공백3개 + 본문"의 형태이다
        # 공백3개는 그냥 공백이 아니고 nbsp이기에 일반공백으로 replace
        # 앞 3문자를 자른다 왜? 절 수가 한자리에서 최대 3자리일것인데 앞 세자리를 자르면 1자리수 절수이든 3자리수 절수이든 문자열 맨 좌측은 공백만 남게되기에
        # 맨 왼쪽 공백 지우기 함수 사용
        for index in range(sec, endSec + 1):
            transformedString = elements[index].text.replace(u'\xa0', u' ')  # nbsp 제거하기  
            transformedString = transformedString[2:len(transformedString)]
            transformedString = transformedString.strip()
            # print(transformedString)
            bibleList.append(transformedString)

        # 한 절씩 깔끔하게 문자열로 들어간 배열 반환환
        return bibleList
