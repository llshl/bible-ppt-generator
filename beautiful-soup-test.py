# import requests
# from bs4 import BeautifulSoup
#
# url = 'https://www.bskorea.or.kr/bible/korbibReadpage.php?version=GAE&book=gen&chap=1&sec=1&cVersion=&fontSize=15px&fontWeight=normal#focus'
#
# response = requests.get(url)
# #tdBible1 > span:nth-child(11) > span
# #tdBible1 > span:nth-child(22)
# #tdBible1 > span:nth-child(10)
# if response.status_code == 200:
#     html = response.text
#     soup = BeautifulSoup(html, 'html.parser')
#     title = soup.select('#tdBible1 > span')
#     print(title)
# else :
#     print(response.status_code)
import requests
from bs4 import BeautifulSoup as bs

page = requests.get("https://www.bskorea.or.kr/bible/korbibReadpage.php?version=GAE&book=2sa&chap=7&sec=3&cVersion"
                    "=&fontSize=15px&fontWeight=normal#focus")
soup = bs(page.text, "html.parser")

elements = soup.select('#tdBible1 > span')

for index, element in enumerate(elements, 1):
        print("{} 번째 게시글의 제목: {}".format(index, element.text))
        # print(element.text)