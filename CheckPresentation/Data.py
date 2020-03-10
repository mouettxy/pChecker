from CheckPresentation import Utils

from globals import DEFAULT_THEME_FONT, DEFAULT_FONT


class Data(Utils.Utils):
    """
    Класс реализовывает получение информации из презентации при помощи модуля python-pptx или API PowerPoint
    """

    def slides_count(self):
        """
        :return: Количество слайдов
        :rtype: int
        """
        return len(self.presentation.slides)

    def slide_by_id(self, slide_id):
        """
        :param slide_id: ID слайда
        :return: Слайд
        :rtype: class
        """
        return self.presentation.slides.get(slide_id)

    def slide_dimensions(self):
        """
        :return: Значение ширины и длины слайдов презентации в пикселях.
        :rtype: tuple of (int, int)
        """
        return self.convert_emu_px(self.presentation.slide_width), self.convert_emu_px(self.presentation.slide_height)

    def font_sizes_by_shape(self, shape, flag):
        """
        Возвращает размер шрифта в shape, если размер шрифта не найден, то получаем размер шрифта по умолчанию.
        ================================================================================================================
        :param shape: shape obj/class презентации
        :param flag: true/false использовать/не использовать проверку наличия размера шрифта
        :return font_sizes: Размеры шрифта в shape
        :rtype tuple of float
        """
        font_sizes = []
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                try:
                    if self.is_not_auto_size(shape):
                        font_sizes.append(run.font.size.pt)
                    else:
                        font_scale = shape.text_frame._bodyPr.normAutofit.fontScale
                        font_sizes.append(run.font.size.pt * font_scale / 100)
                except AttributeError:
                    pass
        if flag:
            if len(font_sizes) == 0:
                font_sizes.append(18.0)
        return font_sizes

    def shapes_by_slide_id(self, slide_id):
        """
        :param slide_id: ID слайда
        :return shapes: Обьекты презентации
        :rtype: list of class
        """
        shapes = []
        for shape in self.slide_by_id(slide_id).shapes:
            shapes.append(shape)
        return shapes

    def font_sizes_by_slide_id(self, slide_id):
        """
        :param slide_id: ID слайда
        :return font_sizes: Все существубщие размеры текста в слайде
        :rtype font_sizes: tuple of float
        """
        font_sizes = []
        for shape in self.shapes_by_slide_id(slide_id):
            if self.is_text(shape):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        try:
                            if self.is_not_auto_size(shape):
                                font_sizes.append(run.font.size.pt)
                            else:
                                font_scale = shape.text_frame._bodyPr.normAutofit.fontScale
                                font_sizes.append(run.font.size.pt * font_scale / 100)
                        except AttributeError:
                            pass
        if len(font_sizes) == 0:
            font_sizes.append(18.0)
        return font_sizes

    def paragraph_runs(self):
        """
        :return runs: "Runs" Параграфов в презентации
        :rtype runs: list of class
        """
        runs = []
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if self.is_text(shape):
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            runs.append(run)
        return runs

    def font_sizes(self):
        """
        :return font_sizes: Размеры текста в каждом слайде
        :rtype font_sizes: dict of (int, list of float)
        """
        font_sizes = {}
        for slide in self.presentation.slides:
            slide_index = int(self.presentation.slides.index(slide) + 1)
            font_sizes.update({slide_index: self.font_sizes_by_slide_id(slide.slide_id)})
        return font_sizes

    @staticmethod
    def typefaces_by_shape(shape):
        """
        Получает названия шрифта по shape презентации
        :param shape: shape презентации
        :type shape: class
        :return: Строку с названием шрифта
        :rtype: string
        """
        typeface = ""
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                try:
                    typeface = run.font.name
                except AttributeError:
                    pass
        if typeface is None:
            typeface = DEFAULT_FONT
        if typeface == "+mn-lt":
            typeface = DEFAULT_THEME_FONT
        return typeface

    def typefaces(self):
        """
        :return typefaces: Уникальные название шрифта из презентации
        :rtype typefaces: set
        """
        typefaces = set()
        for run in self.paragraph_runs():
            try:
                typefaces.add(run.font.name)
            except AttributeError:
                pass
        typefaces = list(typefaces)
        for i in range(len(typefaces)):  # На основе темы презентации по умолчанию
            if typefaces[i] is None and DEFAULT_FONT not in typefaces:
                typefaces[i] = DEFAULT_FONT
            else:
                typefaces[i] = DEFAULT_THEME_FONT
            if typefaces[i] == "+mn-lt":
                typefaces[i] = DEFAULT_THEME_FONT
        return typefaces

    def text(self):
        """
        :return text: Весь текст в презентации
        :rtype text: list of str
        """
        text = []
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if self.is_text(shape):
                    text.append(shape.text)
        return text

    def shape_dimensions(self, shape):
        """
        Возвращает словарь измерений для обьекта презентации.
        :param shape: shape презентации
        :type shape: class
        :return: Словарь изменений для обьекта презентаций содержащий ключи x1, y1, x2, y2, width, height
        :rtype: dict string/int
        """
        width, height = self.convert_emu_px(shape.width), self.convert_emu_px(shape.height)
        left, top = self.convert_emu_px(shape.left), self.convert_emu_px(shape.top)
        x1, y1 = left, top
        x2, y2 = width + x1, height + y1
        return {
            'x1': int(x1),
            'x2': int(x2),
            'y1': int(y1),
            'y2': int(y2),
            'width': int(width),
            'left': int(left),
            'top': int(top),
            'height': int(height),
        }
