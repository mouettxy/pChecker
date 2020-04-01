import csv
import inspect
from pathlib import Path

from MSOCONSTANTS import msoOrientationHorizontal
from PresentationExamLayouts import PresentationExamLayouts as Layouts


class PresentationExamAnalyze(object):
    def __init__(self, presentation, application, utils, images):
        super().__init__()
        self._Presentation = presentation
        self._Application = application
        self._Utils = utils
        self._Images = images

    def __find_layout(self):
        layouts = inspect.getmembers(Layouts, predicate=inspect.isfunction)
        for layout in layouts:
            name = layout[0]
            layout_positions = layout[1](self._Utils.convert_points_px(self._Presentation.PageSetup.SlideWidth),
                                         self._Utils.convert_points_px(self._Presentation.PageSetup.SlideHeight))
            positions = []
            elements, collision = set(), set()
            for Slide in self._Presentation.Slides:
                if Slide.SlideIndex == 2 or Slide.SlideIndex == 3:
                    for Shape in Slide.Shapes:
                        elements.add(Shape.Name)
                        s_dims = self._Utils.get_shape_dimensions(Shape)
                        if self._Utils.is_text(Shape):
                            positions = layout_positions[Slide.SlideIndex]['text_blocks'] + \
                                        layout_positions[Slide.SlideIndex]['title']
                        elif self._Utils.is_image(Shape):
                            positions = layout_positions[Slide.SlideIndex]['images']
                        for position in positions:
                            position = {'left': position[0], 'top': position[1],
                                        'width': position[2], 'height': position[3]}
                            if self._Utils.check_collision_between_shapes(s_dims, position):
                                collision.add(Shape.Name)
            if len(elements) == len(collision):
                return True, name
        return False

    def __analyze_presentation(self):
        analyze = {0: False, 1: False, 2: False, 3: False, 4: False, 5: False, 6: False, 13: True}
        prs, typefaces = self._Presentation, set()
        layout = self.__find_layout()

        if prs.Slides.Count == 3:
            analyze[0] = True
        if ((prs.PageSetup.SlideWidth / prs.PageSetup.SlideHeight) *
            (prs.PageSetup.SlideHeight / prs.PageSetup.SlideWidth)) == 1.0:
            analyze[1] = True
        if prs.PageSetup.SlideOrientation == msoOrientationHorizontal:
            analyze[2] = True
        for Slide in self._Presentation.Slides:
            for Shape in Slide.Shapes:
                if self._Utils.is_text(Shape):
                    typefaces.add(Shape.TextFrame.TextRange.Font.Name)
        if len(typefaces) == 1:
            analyze[3] = True
        if self._Images.compare():
            analyze[4] = True
        if layout:
            analyze[5], analyze[6] = layout
        if self._Images.distorted_images():
            analyze[13] = False
        return analyze

    def __analyze_slide_1(self):
        analyze = {7: False, 8: False, 9: False, 10: False}
        Slide = self._Presentation.Slides(1)
        # collect data
        tmpDimensions, f_sizes, overlaps, counter = None, [], set(), 0
        for Shape in Slide.Shapes:
            dimensions = self._Utils.get_shape_dimensions(Shape)
            if tmpDimensions is None:
                tmpDimensions = dimensions
            else:
                overlaps.add(self._Utils.check_collision_between_shapes(dimensions, tmpDimensions))

            if self._Utils.is_title(Shape):
                if not analyze[7] and not analyze[8]:
                    analyze[7] = True
                elif analyze[7] and not analyze[8]:
                    analyze[8] = True
            elif self._Utils.is_text(Shape):
                counter += 1

            f_sizes.append(Shape.TextFrame.TextRange.Font.Size)

        if not analyze[7] and not analyze[8]:
            if counter == 2:
                analyze[7], analyze[8] = True, True
            if counter == 1:
                analyze[7] = True

        if len(overlaps) == 1:
            analyze[9] = True

        if len(f_sizes) == 2:
            if 40.0 in f_sizes and 24.0 in f_sizes:
                analyze[10] = True

        return analyze

    def __base_analyze(self, slide):
        tmpDimensions, overlaps, f_sizes, text_blocks, image_blocks, title = None, set(), [], 0, 0, False
        Slide = self._Presentation.Slides(slide)
        for Shape in Slide.Shapes:
            dimensions = self._Utils.get_shape_dimensions(Shape)
            if tmpDimensions is None:
                tmpDimensions = dimensions
            else:
                overlaps.add(self._Utils.check_collision_between_shapes(dimensions, tmpDimensions))
            if self._Utils.is_title(Shape):
                if not title:
                    title = True
            elif self._Utils.is_text(Shape):
                text_blocks += 1
                f_sizes.append(Shape.TextFrame.TextRange.Font.Size)
            elif self._Utils.is_image:
                image_blocks += 1
        return title, text_blocks, image_blocks, overlaps, f_sizes

    def __analyze_slide_2(self):
        analyze = {7: False, 9: False, 10: False, 11: False, 12: False}
        analyze[7], text_blocks, image_blocks, overlaps, f_sizes = self.__base_analyze(2)
        if text_blocks == 3 and analyze[7]:
            analyze[7], analyze[11] = True, True
        elif text_blocks == 3 and not analyze[7]:
            analyze[7], analyze[11] = True, True
        elif text_blocks == 2 and not analyze[7]:
            analyze[7], analyze[11] = False, True
        else:
            analyze[7], analyze[11] = False, False
        if len(overlaps) == 1:
            analyze[9] = True
        if f_sizes.count(20.0) == 2 and not analyze[7]:
            analyze[10] = True
        elif f_sizes.count(20.0) == 2 and f_sizes.count(24.0) and analyze[7]:
            analyze[10] = True
        if image_blocks == 2:
            analyze[12] = True
        return analyze

    def __analyze_slide_3(self):
        analyze = {7: False, 9: False, 10: False, 11: False, 12: False}
        analyze[7], text_blocks, image_blocks, overlaps, f_sizes = self.__base_analyze(3)
        if text_blocks == 3 and analyze[7]:
            analyze[7], analyze[11] = True, True
        elif text_blocks == 4 and not analyze[7]:
            analyze[7], analyze[11] = True, True
        elif text_blocks == 3 and not analyze[7]:
            analyze[7], analyze[11] = False, True
        else:
            analyze[7], analyze[11] = False, False
        if len(overlaps) == 1:
            analyze[9] = True
        if f_sizes.count(20.0) == 3 and not analyze[7]:
            analyze[10] = True
        elif f_sizes.count(20.0) == 3 and f_sizes.count(24.0) and analyze[7]:
            analyze[10] = True
        if image_blocks == 3:
            analyze[12] = True
        return analyze

    @staticmethod
    def __translate(analyze, grade=None):
        presentation_analyze = {
            "Соотношение сторон 16:9": "Выполнено" if analyze[1] else "Не выполнено",
            "Горизонтальная ориентация": "Выполнено" if analyze[2] else "Не выполнено"
        }

        structure_analyze = {
            "Три слайда в презентации": "Выполнено" if analyze[0] else "Не выполнено",
            "Соответствует макету": "Выполнено" if analyze[5] else "Не выполнено",
            "Заголовки на слайдах": "Выполнено" if analyze[7] else "Не выполнено",
            "Подзаголовок на первом слайде": "Выполнено" if analyze[8] else "Не выполнено",
            "Элементы не перекрывают друг друга": "Выполнено" if analyze[9] else "Не выполнено",
            "Текстовые блоки на 2, 3 слайде": "Выполнено" if analyze[11] else "Не выполнено",
            "Картинки на 2, 3 слайде": "Выполнено" if analyze[12] else "Не выполнено"
        }

        fonts_analyze = {
            "Единый тип шрифта": "Выполнено" if analyze[3] else "Не выполнено",
            "Размер шрифта": "Выполнено" if analyze[10] else "Не выполнено"
        }

        images_analyze = {
            "Оригинальные картинки": "Выполнено" if analyze[4] else "Не выполнено",
            "Картинки не искажены": "Выполнено" if analyze[13] else "Не выполнено"
        }

        if grade:
            return presentation_analyze, structure_analyze, fonts_analyze, images_analyze, analyze[6], grade
        return presentation_analyze, structure_analyze, fonts_analyze, images_analyze, analyze[6]

    def __summary(self, grade=True):
        """
        0 - does prs have 3 slides
        1 - right aspect ratio of presentation
        2 - horizontal orientation
        3 - right typefaces
        4 - original photos
        5 - contatins layout
        6 - which layout
        7 - title
        8 - subtitle
        9 - overlaps
        10 - font sizes
        11 - text blocks
        12 - image blocks
        13 - distorted images

        Presentation:
            1, 2
        Structure:
            0, 5, 7, 8, 9, 11, 12
        Fonts:
            3, 10
        Images:
            4 13
        """
        # 8 HOURS HERE
        presentation_info, first_slide, second_slide = (self.__analyze_presentation(),
                                                        self.__analyze_slide_1(),
                                                        self.__analyze_slide_2())
        if self._Presentation.Slides.Count >= 3:
            third_slide = self.__analyze_slide_3()
            print(presentation_info)
            print(first_slide)
            print(second_slide)
            print(third_slide)
            data = {**presentation_info, **first_slide, **second_slide, **third_slide}
            err_structure = [data[k] for k in data if k in [0, 5, 7, 8, 9, 11, 12]].count(False)
            err_fonts = [data[k] for k in data if k in [3, 10]].count(False)
            err_images = [data[k] for k in data if k in [4, 13]].count(False)
            if grade:
                r_grade = 0
                # check if we can give grade 2 (max)
                if all(value for value in data.values()):
                    r_grade = 2
                    return self.__translate(data, r_grade)
                # or if we can give 1
                if err_structure == 1 and not err_fonts and not err_images:
                    r_grade = 1
                elif not err_structure and err_fonts == 1 and not err_images:
                    r_grade = 1
                elif not err_structure and not err_fonts and err_images == 1:
                    r_grade = 1
                return self.__translate(data, grade=r_grade)
            self.__translate(data)
        elif self._Presentation.Slides.Count == 2:
            data = {**presentation_info, **first_slide, **second_slide}
            err_structure = [data[k] for k in data if k in [0, 5, 7, 8, 9, 11, 12]].count(False)
            err_fonts = [data[k] for k in data if k in [3, 10]].count(False)
            err_images = [data[k] for k in data if k in [4, 13]].count(False)
            if grade:
                r_grade = 0
                if err_structure == 1 and not err_fonts and not err_images:
                    r_grade = 1
                return self.__translate(data, grade=r_grade)
            return self.__translate(data)

    def get(self, typeof="analyze"):
        if typeof == "analyze":
            return self.__summary()
        elif typeof == "analyze2":
            return self.__summary(grade=False)
        elif typeof == "thumb":
            return self._Images.get("thumb")
        elif typeof == "slides":
            return self._Presentation.Slides.Count

    @property
    def warnings(self):
        warnings = {0: [], 1: [], 2: [], 3: []}
        shape_animations, slide_1_text_blocks = 0, 0
        for Slide in self._Presentation.Slides:
            # count slide animations or entry effects
            if Slide.TimeLine.MainSequence.Count >= 1:
                shape_animations += Slide.TimeLine.MainSequence.Count
            if Slide.SlideShowTransition.EntryEffect:
                warnings[0].append(f"Найдена анимация перехода на слайде номер {Slide.SlideIndex}.")

            for Shape in Slide.Shapes:
                crop = self._Utils.get_shape_crop_values(Shape)
                if crop:
                    crop_warning = (f'Картинка {Shape.Name} с ID {Shape.Id} обрезана слева/справа/сверху/cнизу на '
                                    f'{crop["left"]}/{crop["right"]}/{crop["top"]}/{crop["bottom"]}')
                if Slide.SlideIndex == 1:
                    if self._Utils.is_image(Shape):
                        warnings[1].append(f"Изображение {Shape.Name} с ID {Shape.Id}")
                    elif self._Utils.is_text(Shape) is True:
                        slide_1_text_blocks += 1
                    elif self._Utils.is_text(Shape) is None:
                        warnings[1].append(f"Пустой текстовый блок {Shape.Name} с ID {Shape.Id}")
                    elif not self._Utils.is_text(Shape):
                        warnings[1].append(f"Неопознаный тип обьекта {Shape.Name} с ID {Shape.Id}")
                    if slide_1_text_blocks > 2:
                        warnings[1].append(f"Больше двух текстовых элементов на слайде.")
                elif Slide.SlideIndex == 2:
                    if self._Utils.is_text(Shape) is None:
                        warnings[2].append(f"Пустой текстовый блок {Shape.Name} с ID {Shape.Id}")
                    elif not self._Utils.is_text(Shape) and not self._Utils.is_image(Shape):
                        warnings[2].append(f"Неопознаный тип обьекта {Shape.Name} с ID {Shape.Id}")
                    if crop:
                        warnings[2].append(crop_warning)
                elif Slide.SlideIndex == 3:
                    if self._Utils.is_text(Shape) is None:
                        warnings[3].append(f"Пустой текстовый блок {Shape.Name} с ID {Shape.Id}")
                    elif not self._Utils.is_text(Shape) and not self._Utils.is_image(Shape):
                        warnings[3].append(f"Неопознаный тип обьекта {Shape.Name} с ID {Shape.Id}")
                    if crop:
                        warnings[3].append(crop_warning)
        if shape_animations:
            warnings[0].append(f"Найдены анимации в объектах: {shape_animations}.")
        return warnings

    def export_csv(self):
        warnings = self.warnings
        presentation, structure, fonts, images, layout, grade = self.get()
        warn_0, warn_1, warn_2, warn_3 = warnings[0], warnings[1], warnings[2], warnings[3]
        path = Path.joinpath(Path(self._Utils.get_download_path()), self._Presentation.Name + ".csv")
        fieldnames = ['Презентация', 'Структура', 'Шрифты', 'Картинки', 'Предупреждения', 'Слайд 1', 'Слайд 2',
                      'Слайд 3']
        with open(path, "w", newline='', encoding="windows-1251") as fCsv:
            writer = csv.writer(fCsv, delimiter=',')
            writer.writerow(fieldnames)
            writer.writerow([
                self._Utils.dict_to_string(presentation),
                self._Utils.dict_to_string(structure),
                self._Utils.dict_to_string(fonts),
                self._Utils.dict_to_string(images),
                '\n'.join(warn_0),
                '\n'.join(warn_1),
                '\n'.join(warn_2),
                '\n'.join(warn_3),
            ])
        return path
