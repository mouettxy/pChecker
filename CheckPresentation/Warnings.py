from CheckPresentation import Data

from globals import POSSIBLE_CROP_VALUE_ERROR


class Warnings(Data.Data):
    """
    Класс реализовывает выявление потенциальных ошибок в работе программы, исходя из данных в презентации
    """
    def check_font_sizes(self):
        """
        Проверяет не установлено ли в каком либо текстовом обьекте значение по умолчанию. В таком случае значение
        нельзя получить в данной реализации программы: https://github.com/scanny/python-pptx/issues/337
        :return: False если все значения шрифта корректны, иначе список с некорректными классами.
        :rtype: list of class
        """
        result = []
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if self.is_text(shape):
                    if len(self.font_sizes_by_shape(shape, False)) == 0:
                        result.append(shape)
        if len(result) == 0:
            return False
        return result

    def check_slides(self):
        """
        Проверяет если в презентации больше трёх слайдов.
        :return: True если больше 3х слайдов иначе False
        :rtype: bool
        """
        if self.slides_count() > 3:
            return True
        return False

    def check_crop_image(self):
        """
        Проверяет обрезалось ли изображение. Использует константное значение POSSIBLE_CROP_VALUE_ERROR == 0.05,
        что равно 5% обрезанного. Что укладывается в погрешность.
        :return: Обьект картинки в презенатции
        :rtype: list of class
        """
        result = []
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if self.is_image(shape):
                    if shape.crop_bottom > POSSIBLE_CROP_VALUE_ERROR or \
                            shape.crop_left > POSSIBLE_CROP_VALUE_ERROR or \
                            shape.crop_right > POSSIBLE_CROP_VALUE_ERROR or \
                            shape.crop_top > POSSIBLE_CROP_VALUE_ERROR:
                        result.append(shape)
        if len(result) == 0:
            return False
        return result

    def all_warnings(self):
        """
        Вернёт все предупреждения найденные в презентации
        :return: Строчки предупреждений
        :rtype: list of string
        """
        result = []
        if self.check_font_sizes():
            for shape in self.check_font_sizes():
                result.append(f'Предупреждение: в объекте презентации {shape.name} не удалось получить шрифт. При '
                              f'проверке использовано значение по умолчанию.')
        if self.check_slides():
            result.append('Предупреждение: количество слайдов больше трёх. Проверяются только первые три слайда. '
                          'Проверка может быть некорректной.')
        if self.check_crop_image():
            for shape in self.check_crop_image():
                text = f'Предупреждение: изображение {shape.name} было обрезано пользователем. Детали: \n'
                if shape.crop_bottom:
                    text += f"Снизу обрезано: {shape.crop_bottom * 100}%; \n"
                if shape.crop_top:
                    text += f"Снизу обрезано: {shape.crop_top * 100}%; \n"
                if shape.crop_left:
                    text += f"Снизу обрезано: {shape.crop_left * 100}%; \n"
                if shape.crop_right:
                    text += f"Снизу обрезано: {shape.crop_right * 100}%; \n"
                result.append(text)
        return result