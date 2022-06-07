from pptx import Presentation  # 라이브러리
from pptx.util import Inches  # 사진, 표등을 그리기 위해
from pptx.enum.text import PP_ALIGN  # 정렬 설정하기
from pptx.util import Pt  # Pt 폰트사이즈


class PresentationBuilder:
    pptx_fpath = './220605_청년부예배.pptx'
    presentation = Presentation(pptx_fpath)

    @property
    def xml_slides(self):
        return self.presentation.slides._sldIdLst  # pylint: disable=protected-access

    def move_slide2(self, old_index, new_index):
        slides = list(self.xml_slides)
        self.xml_slides.remove(slides[old_index])
        self.xml_slides.insert(new_index, slides[old_index])

    def move_slide(self, target, indev):
        slides = list(self.xml_slides)
        self.xml_slides.insert(indev, slides[target])

    # also works for deleting slides
    def delete_slide(self, index):
        slides = list(self.xml_slides)
        self.xml_slides.remove(slides[index])

    # also works for deleting slides
    def save_ppt(self):
        self.presentation.save('./220605_청년부예배.pptx')

instance1 = PresentationBuilder()
instance1.move_slide2(4, 5) # 0번(첫번째장) 슬라이드가 46번째 위치(46번째장 위치)로 오게됨(기존 46번째 슬라이드는 45번이 됨)
instance1.save_ppt()
