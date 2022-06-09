from BibleCrawler import BibleCrawler
from PresentationBuilder import PresentationBuilder
from pptx import Presentation  # 라이브러리
from pptx.util import Inches  # 사진, 표등을 그리기 위해
from pptx.enum.text import PP_ALIGN  # 정렬 설정하기
from pptx.util import Pt  # Pt 폰트사이즈

def main():
    # python 3에서는 print() 으로 사용합니다.
    print("Start Program")
    pptx_fpath = './220605_청년부예배.pptx'
    presentation = Presentation(pptx_fpath)
    pptBuilder = PresentationBuilder(presentation)


#여기서부터
    # 피피티 파일 전체 슬라이드 장수
    beforeSize = len(presentation.slides)
    print("before size: " + str(beforeSize))

    # 본문 말씀 절수
    bodyPage = 29 - 22

    # 피피티에서 본문 말씀이 시작되는 위치
    bodyStartPage = 50

    # 본문 말씀 절수만큼 새 슬라이드를 맨뒤에 추가하기
    for i in range(0, bodyPage):
        pptBuilder.add_slide(baseSlidePage=bodyStartPage)

    # 추가한 후 전체 슬라이드 장수
    afterSize = len(presentation.slides)
    print("after size: " + str(afterSize))

    # 추가된 슬라이드들 중 첫번째 슬라이드의 인덱스
    startCopyPageIndex = afterSize - bodyPage
    print(startCopyPageIndex)

    # 추가된 슬라이드들 중 마지막 놈의 인덱스
    lastPageIndex = afterSize - 1
    print(lastPageIndex)

    # 추가된 슬라이드들의 인덱스를 가리키는 포인터
    currentBodyStartPointer = bodyStartPage

    # 추가된 슬라이드들을 포문으로 돌면서 본문 시작위치 뒤로 이동시키기
    for index in range (startCopyPageIndex, lastPageIndex+1):
        print("인덱스 " + str(index) + "번째 슬라이드가 " + str(currentBodyStartPointer) + "번째 인덱스로 이동")
        pptBuilder.move_slide(index, currentBodyStartPointer)
        currentBodyStartPointer += 1

    pptBuilder.save_ppt()
#여기까지가 말씀페이지수만큼 복사해서 중간에 삽입하기

    book = "jhn"
    chap = 14
    sec = 22
    endSec = 29
    bibleCrawler = BibleCrawler()
    bibleList = bibleCrawler.getBible(book, chap, sec, endSec)

    pptBuilder.insert_text(50, 29 - 22, bibleList)
    pptBuilder.save_ppt()

# 결국 최종적으로 bodyStartPage번 인덱스부터 (bodyStartPage + bodyPage)번 인덱스까지 본문내용이다
# book = "jhn"
# chap = 14
# sec = 22
# endSec = 29
# bibleCrawler = BibleCrawler()
# bibleList = bibleCrawler.getBible(book, chap, sec, endSec)
#print(bibleCrawler.getBible(book, chap, sec, endSec))
if __name__ == "__main__":
    main()
