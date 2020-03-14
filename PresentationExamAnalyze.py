from MSOCONSTANTS import msoPlaceholder, msoOrientationHorizontal
from MSOCONSTANTS import ppPlaceholderCenterTitle, ppPlaceholderTitle, ppPlaceholderSubtitle
from PresentationExamLayouts import PresentationExamLayouts as Layouts
import inspect


class PresentationExamAnalyze(object):
    def __init__(self, presentation, application, utils, images):
        super().__init__()
        self._Presentation = presentation
        self._Application = application
        self._Utils = utils
        self._Images = images
        self._warnings = {
            'Предупреждения в первом слайде': [],
            'Предупреждения во втором слайде': [],
            'Предупреждения в третьем слайде': [],
            'Предупреждения по презентации': [],
        }
        self._typefaces = set()

    @property
    def warnings(self):
        return self._warnings

    def __analyze_presentation_slide_parameters(self):
        result = {
            'three_slides': False,
            'aspect_ratio': False,
            'orientation': False,
            'original_photos': False,
        }
        if self._Presentation.PageSetup.SlideOrientation == msoOrientationHorizontal:
            result['orientation'] = True
        if ((self._Presentation.PageSetup.SlideWidth / self._Presentation.PageSetup.SlideHeight) *
            (self._Presentation.PageSetup.SlideHeight / self._Presentation.PageSetup.SlideWidth)) == 1.0:
            result['aspect_ratio'] = True
        if self._Presentation.Slides.Count == 3:
            result['three_slides'] = True

        # check how many animations have presentation and if have generate warning
        # if slide have entry effect generates warning too
        shape_animations = 0
        for Slide in self._Presentation.Slides:
            if Slide.TimeLine.MainSequence.Count >= 1:
                shape_animations += Slide.TimeLine.MainSequence.Count
            if Slide.SlideShowTransition.EntryEffect:
                self._warnings["Предупреждения по презентации"].append(
                    f"Найдена анимация перехода на слайде номер {Slide.SlideIndex}.")
        if shape_animations:
            self._warnings["Предупреждения по презентации"].append(
                f"Найдены анимации в объектах: {shape_animations}.")
        if self._Images.compare():
            result['original_photos'] = True

        return result

    def __analyze_first_slide(self):
        Slide = self._Presentation.Slides(1)
        result = {
            "has_title": False,
            "has_subtitle": False,
            "correct_font_size": False
        }
        font_sizes = []
        reserve_object_counter = 0
        for Shape in Slide.Shapes:
            # generate warning if first shape has images
            if self._Utils.is_image(Shape):
                self._warnings['Предупреждения в первом слайде'].append(f"Изображение {Shape.Name} с ID {Shape.Id}")
                continue
            # process placeholders
            if Shape.Type == msoPlaceholder:
                if Shape.PlaceholderFormat.Type == ppPlaceholderCenterTitle:
                    result['has_title'] = True
                elif Shape.PlaceholderFormat.Type == ppPlaceholderSubtitle:
                    result['has_subtitle'] = True
                elif Shape.PlaceholderFormat.Type == ppPlaceholderTitle:
                    if not result['has_title'] and result['has_subtitle']:
                        result['has_title'] = True
                    elif result['has_title'] and not result['has_subtitle']:
                        result['has_subtitle'] = False
                    elif not result['has_title'] and not result['has_subtitle']:
                        reserve_object_counter += 1
                else:
                    # uses reserve because object can be image or text or placeholder
                    # it's not placeholder, it's not image, that can be any object that PowerPoint have
                    # generating warning in this case, and counts this as text
                    reserve_object_counter += 1
                    self._warnings['Предупреждения в первом слайде'].append(
                        f"Непредсказанный тип объекта {Shape.Name} с ID {Shape.Id}"
                    )
            # reserve algorithm if no placeholders in slide, we count whole objects
            else:
                if self._Utils.is_text(Shape):
                    reserve_object_counter += 1
            # process font sizes
            font_sizes.append(Shape.TextFrame.TextRange.Font.Size)
            # add typefaces for future use
            self._typefaces.add(Shape.TextFrame.TextRange.Font.Name)
        # last chance to get correct slide, if no one placeholders found, has_title and has_subtitle be False
        # we just counting text objects
        if not result['has_title'] and not result['has_subtitle']:
            if reserve_object_counter == 1:
                result['has_title'] = True
            elif reserve_object_counter == 2:
                result['has_title'] = True
                result['has_subtitle'] = True
        # if font size contains 2 elements, there 2 text elements in slide, that correct, else - generate warning
        # but process anyway
        if len(font_sizes) == 2:
            if 40.0 in font_sizes and 24.0 in font_sizes:
                result['correct_font_size'] = True
        else:
            self._warnings['Предупреждения в первом слайде'].append(f"Больше двух текстовых элементов на слайде.")
        return result

    def __analyze_second_slide(self):
        pass

    def __analyze_third_slide(self):
        pass

    def __summary(self):
        pass

    @property
    def analyze(self):
        return self.__summary()
