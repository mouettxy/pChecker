import base64
import os

from CheckPresentation import Data

from globals import POSSIBLE_DISTORTED_IMAGE_VALUE
from collections import Counter

from PIL import Image


class Analyze(Data.Data):
    """
    Класс реализовывает анализ полученых данных из полученной презентации.
    """
    def most_common_words(self):
        """
        :return most_common: Топ 5 слов в презентации
        :rtype most_common: list of tuple of (str, int)
        """
        text_analyzed = []
        for text in self.text():
            text_analyzed.extend(self.string_optimize(text).split())
        most_common = Counter(text_analyzed).most_common(5)
        return most_common

    def image_overlaps(self):
        """
        Проверяет обьекты презентации на коллизии.
        ================================================================================================================
                                                                WIP
        ================================================================================================================
        :return: Список с строками вида "Коллизия между: обьект1 и обьект2 на номер_слайда слайде
        :rtype: list of string
        """
        result = []
        for slide in self.presentation.slides:
            index = int(self.presentation.slides.index(slide) + 1)
            previous_dimensions = []
            previous_shape = None
            for shape in slide.shapes:
                dim = self.shape_dimensions(shape)
                current_dimensions = [dim['x1'], dim['y1'], dim['width'], dim['height']]
                if len(previous_dimensions) > 0 and previous_shape is not None:
                    if self.check_collision(current_dimensions, previous_dimensions):
                        result.append(f"Коллизия между: {shape.name} и {previous_shape.name} на {index} слайде")
                previous_dimensions = current_dimensions
                previous_shape = shape
        return result

    def distorted_images(self):
        """
        Проверяет пропорциональность изображений учитывая возожную погрешность до 20 пикселей в каждую сторону.
        :return: True если хоть одна картинка не пропорциональна, False иначе
        :return: boolean
        """
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if self.is_image(shape):
                    tmp = open('tmp.jpg', 'wb')
                    tmp.write(base64.b64decode(base64.b64encode(shape.image.blob)))
                    tmp.close()
                    tmp_pil = Image.open('tmp.jpg')
                    # ширина и высота картинки в презентации
                    p_w, p_h = self.convert_emu_px(shape.width), self.convert_emu_px(shape.height)
                    # ширина и высота оригинальной картинки
                    o_w, o_h = tmp_pil.size
                    tmp_pil.close()
                    os.remove('tmp.jpg')
                    need_width, need_height = o_w * (p_h / o_h), o_h * (p_w / o_w)
                    diff_width, diff_height = abs(p_w - need_width), abs(p_h - need_height)
                    if diff_width > POSSIBLE_DISTORTED_IMAGE_VALUE or diff_height > POSSIBLE_DISTORTED_IMAGE_VALUE:
                        return True
        return False

    def slides_contents(self):
        """
        :return slides: Возвращает количество блоков текста, изображений, названий, в каждом слайде с 1-3 включительно.
        """
        slides = {
            1: {
                'textCounter': 0,
                'pictureCounter': 0,
                'titleCounter': 0,
            },
            2: {
                'textCounter': 0,
                'pictureCounter': 0,
                'titleCounter': 0,
            },
            3: {
                'textCounter': 0,
                'pictureCounter': 0,
                'titleCounter': 0,
            }
        }
        for slide in self.presentation.slides:
            index = int(self.presentation.slides.index(slide) + 1)
            for shape in slide.shapes:
                if self.is_text(shape):
                    if self.is_title(shape):
                        slides[index]['titleCounter'] += 1
                    else:
                        slides[index]['textCounter'] += 1
                if self.is_image(shape):
                    slides[index]['pictureCounter'] += 1
        return slides

    def analyze(self):
        """
        :return: Dict где ключ => ID критерия, значение => Int or Boolean
        .. note:: ID критериев:
            Int 0 => Str 'Количество слайдов'
            Int 1 => Str 'Блоки текста и изображений размещены'
            Int 2 => Str 'Название на титульном'
            Int 3 => Str 'Название на 2м и 3м слайде'
            Int 4 => Str 'Единый шрифт'
            Int 5 => Str 'Правильный размер шрифта'
            Int 6 => Str 'Изображения не искажены'
        """
        slides_count = self.slides_count()
        analyze_params = {
            0: slides_count,
            1: None,
            2: None,
            3: None,
            4: None,
            5: None,
            6: self.distorted_images(),
        }
        slides_contents = self.slides_contents()
        font_sizes = self.font_sizes()
        typefaces = self.typefaces()
        # Check first slide #
        slide1_font_size = 0
        slide1_blocks_correct = 0
        if max(font_sizes[1]) == 40 and min(font_sizes[1]) == 24:
            slide1_font_size = 1
        if slides_contents[1]['titleCounter'] + slides_contents[1]['textCounter'] == 2:
            slide1_blocks_correct = 1
        if slides_contents[1]['titleCounter'] >= 1 or \
                (slides_contents[1]['textCounter'] >= 1 and not slides_contents[1]['titleCounter']):
            analyze_params[2] = True
        # End first slide #

        # Check second slide #
        slide2_font_size = 0
        slide2_blocks_correct = 0
        slide2_title = 0
        if max(font_sizes[2]) == 24 and min(font_sizes[2]) == 20:
            slide2_font_size = 1
        if slides_contents[2]['textCounter'] + slides_contents[2]['titleCounter'] == 3 and \
                slides_contents[2]['pictureCounter'] == 2:
            slide2_blocks_correct = 1
            slide2_title = 1
        if slides_contents[2]['textCounter'] + slides_contents[2]['titleCounter'] == 2 and \
                slides_contents[2]['pictureCounter'] == 2:
            slide2_blocks_correct = 1
        # End second slide #

        # Check third slide #
        slide3_font_size = 0
        slide3_blocks_correct = 0
        slide3_title = 0
        if max(font_sizes[3]) == 24 and min(font_sizes[3]) == 20:
            slide3_font_size = 1
        if slides_contents[3]['textCounter'] + slides_contents[3]['titleCounter'] == 3 and \
                slides_contents[3]['pictureCounter'] == 3:
            slide3_blocks_correct = 1
        if slides_contents[3]['textCounter'] + slides_contents[3]['titleCounter'] == 4 and \
                slides_contents[3]['pictureCounter'] == 3:
            slide3_blocks_correct = 1
            slide3_title = 1
        if slides_contents[3]['titleCounter'] == 1:
            slide3_title = 1
        # End third slide #

        if (slide1_blocks_correct + slide2_blocks_correct + slide3_blocks_correct) == 3:
            analyze_params[1] = True
        else:
            analyze_params[1] = False
        if (slide2_title + slide3_title) == 2:
            analyze_params[3] = True
        else:
            analyze_params[3] = False
        if (slide1_font_size + slide2_font_size + slide3_font_size) == 3:
            analyze_params[6] = True
        else:
            analyze_params[6] = False
        if not len(typefaces) > 1:
            analyze_params[5] = True
        else:
            analyze_params[5] = False
        return self._translate_results(analyze_params)
