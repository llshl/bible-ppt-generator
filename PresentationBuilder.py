import copy

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

