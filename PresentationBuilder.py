import copy
from pptx import Presentation  # 라이브러리
from pptx.util import Inches  # 사진, 표등을 그리기 위해
from pptx.enum.text import PP_ALIGN  # 정렬 설정하기
from pptx.util import Pt  # Pt 폰트사이즈
from pptx.dml.color import RGBColor

class PresentationBuilder:
    def __init__(self, resourcePpt):
        self.presentation = resourcePpt

    @property
    def xml_slides(self):
        return self.presentation.slides._sldIdLst  # pylint: disable=protected-access

    # def move_slide2(self, old_index, new_index):
    #     slides = list(self.xml_slides)
    #     self.xml_slides.remove(slides[old_index])
    #     self.xml_slides.insert(new_index, slides[old_index])

    def add_slide(self, baseSlidePage):
        slide_layout = self.presentation.slide_layouts[6]  # 빈 슬라이드 선언
        copy_slide = self.presentation.slides.add_slide(slide_layout)
        baseSlide = self.presentation.slides[baseSlidePage]
        # 기존 슬라이드에 있던 shape들 복사
        for shape in baseSlide.shapes:
            el = shape.element
            newel = copy.deepcopy(el)
            copy_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')

    def move_slide(self, target, index):
        slides = list(self.xml_slides)
        self.xml_slides.insert(index, slides[target])

    # also works for deleting slides
    def delete_slide(self, index):
        slides = list(self.xml_slides)
        self.xml_slides.remove(slides[index])

    # also works for deleting slides
    def save_ppt(self):
        self.presentation.save('./220605_청년부예배.pptx')

    def insert_text(self, start_slide_num, count, bibleList):
        print("start_slide_num은 {}이고 count는 {}이다".format(start_slide_num, count))
        index = 0
        print("index는 {}부터 시작이고".format(index))
        print("성경구절:")
        print(bibleList)
        for i in range(start_slide_num, start_slide_num + count + 1):
            slide_num = start_slide_num + index
            print("{}번째 반복이다. {}인덱스 슬라이드에 작업중이다".format(index + 1, slide_num))
            slide = self.presentation.slides[slide_num]
            shapes_list = slide.shapes
            print("here")
            shape_index = {}
            for i, shape in enumerate(shapes_list):
                shape_index[shape.name] = i
            shape_name = "본문내용"
            print("plz")
            print(shapes_list)
            print(shape_index)
            shape_select = shapes_list[shape_index[shape_name]]

            # shape 내 텍스트 프레임 선택하기 & 기존 값 삭제하기
            text_frame = shape_select.text_frame
            text_frame.clear()

            # 문단 선택하기
            p = text_frame.paragraphs[0]

            # 정렬 설정 : 중간정렬
            p.alighnment = PP_ALIGN.CENTER

            # 텍스트 입력 / 폰트 지정
            run = p.add_run()
            run.text = bibleList[index]
            font = run.font
            font.name = "맑은 고딕"
            font.size = Pt(28)
            font.color.rgb = RGBColor(112, 148, 138)
            font.bold = True
            # font.name = None  # 지정하지 않으면 기본 글씨체로  #  'Arial'

            index += 1