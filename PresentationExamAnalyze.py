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
        self._layouts = inspect.getmembers(Layouts, predicate=inspect.isfunction)
        self._warnings = {
            'Предупреждения в первом слайде': [],
            'Предупреждения во втором слайде': [],
            'Предупреждения в третьем слайде': [],
            'Предупреждения по презентации': [],
        }
        self._typefaces = set()

    def __analyze_presentation_slide_parameters(self):
        result = {
            'three_slides': False,
            'aspect_ratio': False,
            'orientation': False,
            'typefaces': False,
            'original_photos': False,
            'contains_layout': False,
            'which_layout': None,
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

        # layout check
        for layout in self._layouts:
            layout_name = layout[0]
            layout_dimensions = layout[1](self._Utils.convert_points_px(self._Presentation.PageSetup.SlideWidth),
                                          self._Utils.convert_points_px(self._Presentation.PageSetup.SlideHeight))
            rectangles = []
            all_objects, objects_with_collision = set(), set()
            for Slide in self._Presentation.Slides:
                if Slide.SlideIndex == 2 or Slide.SlideIndex == 3:
                    for Shape in Slide.Shapes:
                        all_objects.add(Shape.Name)
                        shape_dimensions = self._Utils.get_shape_dimensions(Shape)
                        if self._Utils.is_text(Shape):
                            rectangles = layout_dimensions[Slide.SlideIndex]['text_blocks']
                        elif self._Utils.is_image(Shape):
                            rectangles = layout_dimensions[Slide.SlideIndex]['images']
                        for r in rectangles:
                            rectangle_dimensions = {'left': r[0], 'top': r[1], 'width': r[2], 'height': r[3]}
                            if self._Utils.check_collision_between_shapes(shape_dimensions,
                                                                          rectangle_dimensions):
                                objects_with_collision.add(Shape.Name)
            if len(all_objects) == len(objects_with_collision):
                result['contains_layout'] = True
                result['which_layout'] = layout_name

        # typefaces
        for Slide in self._Presentation.Slides:
            for Shape in Slide.Shapes:
                if self._Utils.is_text(Shape):
                    self._typefaces.add(Shape.TextFrame.TextRange.Font.Name)

        if len(self._typefaces) == 1:
            result['typefaces'] = True

        return result

    def __analyze_first_slide(self):
        Slide = self._Presentation.Slides(1)
        result = {
            "has_title": False,
            "has_subtitle": False,
            "shapes_overlaps": True,
            "correct_font_size": False
        }
        font_sizes, reserve_object_counter, previous_shape, shape_overlaps = [], 0, None, set()
        for Shape in Slide.Shapes:
            shape_dimensions = self._Utils.get_shape_dimensions(Shape)
            if previous_shape is None:
                previous_shape = shape_dimensions
            elif previous_shape is not None:
                shape_overlaps.add(self._Utils.check_collision_between_shapes(shape_dimensions, previous_shape))
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

        if not len(shape_overlaps) > 1:
            result['shapes_overlaps'] = False

        return result

    def __base_slide_analyze(self, slide):
        title, text_blocks, images, font_sizes, previous_shape, shape_overlaps = False, 0, 0, [], None, set()
        for Shape in self._Presentation.Slides(slide).Shapes:
            shape_dimensions = self._Utils.get_shape_dimensions(Shape)
            if previous_shape is None:
                previous_shape = shape_dimensions
            elif previous_shape is not None:
                shape_overlaps.add(self._Utils.check_collision_between_shapes(shape_dimensions, previous_shape))
            if self._Utils.is_text(Shape) is True:
                if self._Utils.is_title(Shape):
                    if not title:
                        title = True
                font_sizes.append(Shape.TextFrame.TextRange.Font.Size)
                text_blocks += 1
            elif self._Utils.is_text(Shape) is None:
                self._warnings['Предупреждения во втором слайде'].append(
                    f"Пустой текстовый блок {Shape.Name} с ID {Shape.Id}"
                )
            else:
                if self._Utils.is_image(Shape):
                    images += 1
        return title, text_blocks, images, font_sizes, shape_overlaps

    def __analyze_second_slide(self):
        result = {
            "has_title": False,
            "has_text_blocks": False,
            "has_images": False,
            "shapes_overlaps": True,
            "correct_font_size": False,
        }

        result['has_title'], text_blocks, images, font_sizes, shape_overlaps = self.__base_slide_analyze(2)

        if text_blocks == 2 and result['has_title']:
            result['has_text_blocks'] = True
        elif text_blocks == 3 and not result['has_title']:
            result['has_text_blocks'] = True
        elif text_blocks == 2 and not result['has_title']:
            result['has_text_blocks'] = True
            # play it safe
            result['has_title'] = False
        else:
            result['has_text_blocks'], result['has_title'] = False, False

        if images == 2:
            result['has_images'] = True

        if 20.0 in font_sizes and not result['has_title']:
            result['correct_font_size'] = True
        elif 20.0 in font_sizes and 24.0 in font_sizes and result['has_title']:
            result['correct_font_size'] = True

        if not len(shape_overlaps) > 1:
            result['shapes_overlaps'] = False

        return result

    def __analyze_third_slide(self):
        result = {
            "has_title": False,
            "has_text_blocks": False,
            "has_images": False,
            "shapes_overlaps": True,
            "correct_font_size": False,
        }

        result['has_title'], text_blocks, images, font_sizes, shape_overlaps = self.__base_slide_analyze(3)

        if text_blocks == 3 and result['has_title']:
            result['has_text_blocks'] = True
        elif text_blocks == 4 and not result['has_title']:
            result['has_text_blocks'] = True
        elif text_blocks == 3 and not result['has_title']:
            result['has_text_blocks'] = True
            # play it safe
            result['has_title'] = False
        else:
            result['has_text_blocks'], result['has_title'] = False, False

        if images == 3:
            result['has_images'] = True

        if 20.0 in font_sizes and not result['has_title']:
            result['correct_font_size'] = True
        elif 20.0 in font_sizes and 24.0 in font_sizes and result['has_title']:
            result['correct_font_size'] = True

        if not len(shape_overlaps) > 1:
            result['shapes_overlaps'] = False

        return result

    def __summary(self, how="detail"):
        """
        :param how: detail / minimal / errors
        :type how: string
        :return:
        """
        data = {
            'average': self.__analyze_presentation_slide_parameters(),
            'first': self.__analyze_first_slide(),
            'second': self.__analyze_second_slide(),
        }
        detail_result = {
            "Презентация": {
                "Слайды 16:9": "Не выполнено.",
                "Горизонтальная ориентация": "Не выполнено.",
            },
            "Структура": {
                "Презентация состоит ровно из трёх слайдов": "Не выполнено.",
                "Информация на слайдах размещена согласно макету": "Не выполнено.",
                "2 и 3 слайды имеют заголовки": "Не выполнено.",
                "Элементы презентации не перекрывают друг друга": "Не выполнено.",
            },
            "Шрифт": {
                "Единый тип шрифта": "Не выполнено.",
                "Правильный размер шрифта": "Не выполнено.",
            },
            "Изображения": {
                "Сохранены пропорции при масштабировании": "Не выполнено.",
                "Соответствуют данным в задании изображениям": "Не выполнено.",
            }
        }

        if data['average']['three_slides']:
            data['third'] = self.__analyze_third_slide()

        if how == "detail":
            detail_result['Структура']['Презентация состоит ровно из трёх слайдов'] = "Выполнено."
            if data['average']['aspect_ratio']:
                detail_result['Презентация']['Горизонтальная ориентация'] = "Выполнено."
            if data['average']['orientation']:
                detail_result['Презентация']['Слайды 16:9'] = "Выполнено."
            if data['average']['contains_layout']:
                detail_result['Структура']['Информация на слайдах размещена согласно макету'] = "Выполнено."
            if data['average']['typefaces']:
                detail_result['Шрифт']['Единый тип шрифта'] = "Выполнено."
            # saved for future use
            # if data['average']['images_aspect_ratio']:
            #     result['Изображения']['Сохранены пропорции при масштабировании'] = "Выполнено."
            if data['average']['original_photos']:
                detail_result['Изображения']['Соответствуют данным в задании изображениям'] = "Выполнено."

            if data['average']['three_slides']:
                detail_result['Структура']['Презентация состоит ровно из трёх слайдов'] = "Выполнено."
                if data['second']['has_title'] and data['third']['has_title']:
                    detail_result['Структура']['2 и 3 слайды имеют заголовки'] = "Выполнено."
                if (not data['first']['shapes_overlaps'] or
                        not data['second']['shapes_overlaps'] or
                        not data['third']['shapes_overlaps']):
                    detail_result['Структура']['Элементы презентации не перекрывают друг друга'] = "Выполнено."
                if (data['first']['correct_font_size'] and
                        data['second']['correct_font_size'] and
                        data['third']['correct_font_size']):
                    detail_result['Шрифт']['Правильный размер шрифта'] = "Выполнено."
            else:
                detail_result['Структура']['Презентация состоит ровно из трёх слайдов'] = "Не выполнено."
                if data['second']['has_title']:
                    detail_result['Структура']['2 и 3 слайды имеют заголовки'] = "Выполнено."
                if (not data['first']['shapes_overlaps'] or
                        not data['second']['shapes_overlaps']):
                    detail_result['Структура']['Элементы презентации не перекрывают друг друга'] = "Выполнено."
                if (data['first']['correct_font_size'] and
                        data['second']['correct_font_size']):
                    detail_result['Шрифт']['Правильный размер шрифта'] = "Выполнено."

            # process grade
            result_grade = 0
            # if all criteria is true we give max grade and leave
            if list(self._Utils.dict_to_list(detail_result)).count("Не выполнено.") == 0:
                result_grade = 2
                return detail_result, result_grade

            # check if we can give grade 1
            structure_c = list(self._Utils.dict_to_list(detail_result, "Структура")).count("Не выполнено.")
            font_c = list(self._Utils.dict_to_list(detail_result, "Шрифт")).count("Не выполнено.")
            images_c = list(self._Utils.dict_to_list(detail_result, "Изображения")).count("Не выполнено.")
            if data['average']['three_slides']:
                if structure_c == 1 and font_c == 0 and images_c == 0:
                    result_grade = 1
                elif structure_c == 0 and font_c == 1 and images_c == 1:
                    result_grade = 1
                elif structure_c == 0 and font_c == 0 and images_c == 1:
                    result_grade = 1
                return detail_result, result_grade
            else:
                if (self._Presentation.Slides.Count == 2 and
                        structure_c == 0 and font_c == 0 and images_c == 0 and
                        data['average']['contains_layout']):
                    result_grade = 1
            return detail_result, result_grade
        elif how == "minimal":
            pass
        elif how == "errors":
            pass
        else:
            return "Не заявленный метод получения результата."

    @property
    def warnings(self):
        return self._warnings

    def get(self, how="detail"):
        return self.__summary(how=how)
