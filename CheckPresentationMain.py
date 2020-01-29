# -*- coding: utf-8 -*-
import re
import json
import os
import pandas as pd

from pathlib import Path
from pandas.errors import EmptyDataError
from collections import Counter
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER_TYPE
from pptx.enum.text import MSO_AUTO_SIZE
from globals import TEXT_THRESHOLD


class CheckPresentationUtils:
    """
    :param self.text_threshold: После скольки символов текст считается осмысленным и его можно засчитывать как текст.
    :type self.text_threshold: int
    """
    def __init__(self):
        super().__init__()
        self.text_threshold = TEXT_THRESHOLD

    @staticmethod
    def _to_json(string):
        """
        :param string: Любая информация
        :type string: dict, list, tuple, str, int, long, float, True, False, None
        :return: JSON формат
        :rtype: object, array, string, number, true, false, null
        """
        return json.dumps(string)

    @staticmethod
    def is_image(shape):
        """
        :param shape: Обьект презентации
        :return: Является ли обьект картинкой
        :rtype: bool
        """
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            return True
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.PICTURE:
            return True
        return False

    @staticmethod
    def is_title(shape):
        """
        :param shape: Обьект презентации
        :return: Является ли обьект заголовком
        :rtype: bool
        """
        if shape.is_placeholder and (
            shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.TITLE
                or shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.SUBTITLE
                or shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.VERTICAL_TITLE
                or shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.CENTER_TITLE):
            return True
        return False

    def is_text(self, shape):
        """
        :param shape: Обьект презентации
        :return: Является ли обьект текстом
        :rtype: bool
        """
        if shape.has_text_frame:
            if self.is_title(shape):
                return True
            if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.BODY:
                return True
            if len(shape.text) > self.text_threshold:
                return True
        return False

    @staticmethod
    def string_optimize(string):
        """
        :param string: Строчка/Параграф из презентации
        :type string: str
        :return: Очищенная строка без лишних символов и слов менее 4х букв
        :rtype: str
        """
        delete_junk_symbols = re.compile('[^a-zA-Zа-яА-ЯёЁ ]')
        delete_junk_words = re.compile('\\b\\w{0,3}\\b')
        return delete_junk_words.sub("", delete_junk_symbols.sub("", string.lower()))


class CheckPresentationGetData(CheckPresentationUtils):
    """
    :param presentation: Обьект Presentation() из python-pptx
    """
    def __init__(self, presentation):
        super().__init__()
        self.presentation  = presentation

    def get_slides_length(self):
        """
        :return: Количество слайдов
        :rtype: int
        """
        return len(self.presentation.slides)

    def get_slides(self):
        """
        :return: Все слайды в презентации
        :rtype: list of class
        """
        return [slide for slide in self.presentation.slides]

    def get_slide_by_id(self, slide_id):
        """
        :param slide_id: ID слайда
        :return: Слайд
        :rtype: class
        """
        return self.presentation.slides.get(slide_id)

    def get_slides_shapes(self):
        """
        :return shapes_by_slide: Обьекты презентации по слайдам
        :rtype shapes_by_slide: dict of (int, list of class)
        """
        shapes_by_slide = {
            1: [],
            2: [],
            3: []
        }
        for slide in self.presentation.slides:
            slide_index = int(self.presentation.slides.index(slide) + 1)
            for shape in slide.shapes:
                shapes_by_slide[slide_index].append(shape)
        return shapes_by_slide

    def get_shapes_by_slide_id(self, slide_id):
        """
        :param slide_id: ID слайда
        :return shapes: Обьекты презентации
        :rtype: list of class
        """
        shapes = []
        for shape in self.get_slide_by_id(slide_id).shapes:
            shapes.append(shape)
        return shapes

    def get_font_sizes_by_slide_id(self, slide_id):
        """
        :param slide_id: ID слайда
        :return font_sizes: Все существубщие размеры текста в слайде
        :rtype font_sizes: list of float
        """
        font_sizes = []
        for shape in self.get_shapes_by_slide_id(slide_id):
            if self.is_text(shape):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        try:
                            if shape.text_frame.auto_size is None or \
                                    shape.text_frame.auto_size == MSO_AUTO_SIZE.NONE or \
                                    shape.text_frame.auto_size == MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT:
                                font_sizes.append(run.font.size.pt)
                            else:
                                font_sizes.append(run.font.size.pt *
                                                         shape.text_frame._bodyPr.normAutofit.fontScale / 100)
                        except AttributeError:
                            pass
        return font_sizes

    def get_all_paragraph_runs(self):
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

    def get_font_sizes(self):
        """
        :return font_sizes: Размеры текста в каждом слайде
        :rtype font_sizes: dict of (int, list of float)
        """
        font_sizes = {}
        for slide in self.presentation.slides:
            slide_index = int(self.presentation.slides.index(slide) + 1)
            font_sizes.update({slide_index: self.get_font_sizes_by_slide_id(slide.slide_id)})
        return font_sizes

    def get_typefaces(self):
        """
        :return typefaces: Уникальные название шрифта из презентации
        :rtype typefaces: set
        """
        typefaces = set()
        for run in self.get_all_paragraph_runs():
            try:
                typefaces.add(run.font.name)
            except AttributeError:
                pass
        return typefaces

    def get_text(self):
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

    def analyze_text(self):
        """
        :return most_common: Топ 5 слов в презентации
        :rtype most_common: list of tuple of (str, int)
        """
        text_analyzed = []
        for text in self.get_text():
            text_analyzed.extend(self.string_optimize(text).split())
        most_common = Counter(text_analyzed).most_common(5)
        print(most_common)
        return most_common


class CheckPresentationAnalyze(CheckPresentationGetData):
    def get_slides_contents(self):
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

    def _translate_results(self, results):
        """
        :param results: Принимает результат работы analyze_results()
        :return: Dict где ключ переведён из ID в Str
        """
        results['Количество слайдов']                   = results.pop(0)
        results['Блоки текста и изображений размещены'] = results.pop(1)
        results['Название на титульном']                = results.pop(2)
        results['Название на 2м и 3м слайде']           = results.pop(3)
        results['Соответствие теме']                    = results.pop(4)
        results['Единый шрифт']                         = results.pop(5)
        results['Правильный размер шрифта']             = results.pop(6)
        results['Текст не перекрывает изображения']     = results.pop(7)
        results['Изображения не искажены']              = results.pop(8)
        results['Изображения не перекрывают элементы']  = results.pop(9)
        return results

    def analyze_results(self, txt_img_collisions_btn=False, distorted_images_btn=False, all_collisions_btn=False,
                        content_compliance=False):
        """
        :param txt_img_collisions_btn: Boolean значение критерия 'Текст не перекрывает изображения'
        :param distorted_images_btn: Boolean значение критерия 'Изображения не искажены'
        :param all_collisions_btn: Boolean значение критерия 'Изображения не перекрывают элементы'
        :param content_compliance: Boolean значение критерия 'Соответствие теме'
        :return: Dict где ключ => ID критерия, значение => Int or Boolean
        .. note:: ID критериев:
            Int 0 => Str 'Количество слайдов'
            Int 1 => Str 'Блоки текста и изображений размещены'
            Int 2 => Str 'Название на титульном'
            Int 3 => Str 'Название на 2м и 3м слайде'
            Int 4 => Str 'Соответствие теме'
            Int 5 => Str 'Единый шрифт'
            Int 6 => Str 'Правильный размер шрифта'
            Int 7 => Str 'Текст не перекрывает изображения'
            Int 8 => Str 'Изображения не искажены'
            Int 9 => Str 'Изображения не перекрывают элементы'
        """
        slides_count = self.get_slides_length()
        if slides_count < 3 or slides_count > 4:
            error = {
                'Ошибка': 'Количество слайдов менее или более трёх.',
            }
            return error

        analyze_params = {
            0: slides_count,
            1: None,
            2: None,
            3: None,
            4: content_compliance,
            5: None,
            6: None,
            7: txt_img_collisions_btn,
            8: distorted_images_btn,
            9: all_collisions_btn,
        }
        slides_contents = self.get_slides_contents()
        font_sizes      = self.get_font_sizes()
        typefaces       = self.get_typefaces()
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


class PrintTo:
    """
    Класс PrintTo
    Конструктор принимает:
    :param results: Переведённый результат работы  CheckPresentationAnalyze.analyze_rezults()
    :param path_to_output: Полный путь к файлу куда будет выгружен результат
    :param path_to_pptx: Полный путь к файлу презентации
    :param encoding: Кодировка которая должна получиться в выходном файле
    TODO: Добавить поддержку изменения кодировки к PrintTo.txt
    TODO: Добавить генерацию файла Excel PrintTo.excel
    """
    def __init__(self, results, path_to_output, path_to_pptx, encoding='utf-8'):
        self.results        = results
        self.path_to_output = Path(path_to_output)
        self.path_to_pptx   = Path(path_to_pptx)
        self.encoding       = encoding
        self.output_name, self.output_extension = os.path.splitext(self.path_to_output)
        self.results_keys   = []
        self.results_values = []
        for result in self.results:
            self.results_keys.append(result)
            self.results_values.append(self.results[result])
        self.results_zip   = list(zip(self.results_keys, self.results_values))
        self.mode_list = ['write', 'rewrite']

    def _extension_check(self, extension):
        """
        :param extension: Расширение файла str (пример: ".txt")
        :return: None в случае когда расширение совпадает с вызванное функцией, иначе генерирует подробную ошибку
        """
        if self.output_extension != extension:
            raise Exception('Ошибка! Неверное расширение файла. \n'
                            f'Ожидалось "{extension}", введено "{self.output_extension}". \n'
                            'Проверьте правильность введённого пути к конечному файлу.')
        return

    def _write_mode_check(self, mode):
        """
        :param mode: Режим записи в файл из значений self.mode_list
        :return: Возвращает str применимую к режиму записи файла
        """
        mode_string = ', '.join(self.mode_list)
        if mode == 'write':
            return 'a'
        elif mode == 'rewrite':
            return 'w'
        else:
            raise Exception('Ошибка! Неверно указан метод открытия файла.\n'
                            f'Ожидаемые значения "{mode_string}", введено "{mode}". \n'
                            'Доступные методы: \n'
                            ' write - файл не будет перезаписан, полученный результат добавиться к предыдущему. \n'
                            ' rewrite - файл будет перезаписан, данные которые были до этого будут удалены. \n')

    def _empty(self, file_path, extension):
        """
        :param file_path: Путь к файлу (пример: 'C:/path/to/file.extension')
        :param extension: Расширение файла (пример: '.txt')
        :return: True если файл пуст, False если файл не пуст
        """
        if extension == '.txt':
            if os.stat(file_path).st_size > 0:
                return False
            return True
        elif extension == '.csv':
            try:
                file_contents = pd.read_csv(file_path)
                return file_contents.empty
            except EmptyDataError:
                return True
            except (OSError, IOError):
                create_file = open(file_path, 'w')
                create_file.close()
                return True
        elif extension == '.xlsx':
            try:
                file_contents = pd.read_excel(file_path)
                return file_contents.empty
            except EmptyDataError:
                return True
            except (OSError, IOError):
                create_file = open(file_path, 'w')
                create_file.close()
                return True

    def _write_to_csv(self, data, file_path, mode, encoding):
        """
        :param data: Список для заполнения
        :param file_path: Путьк файлу
        :param mode: Режим записи в файл
        :param encoding: Кодировка текста внутри
        :return: None или Str в случае выполнения, игаче генерирует подробные ошибки.
        """
        data_frame = pd.DataFrame(data)
        try:
            return data_frame.to_csv(file_path, mode=mode, header=True, encoding=encoding, index=False)
        except PermissionError:
            raise Exception('Недостаточно прав для открытия, создания, или записи в файл. \n'
                            'Попробуйте закрыть файл, или проверить права на запись и чтение файла. \n')

    def txt(self, mode):
        """
        :param mode: Режим записи в файл
        :return: Возвращает str в случае успешного выполнения, иначе генерирует подробные ошибки
        """
        self._extension_check('.txt')
        mode = self._write_mode_check(mode)
        txt_file = open(self.path_to_output, mode=mode, encoding='utf-8')
        if not(self._empty(self.path_to_output, '.txt')):
            txt_file.write('\n')
        txt_file.write(f'Проверка файла: {self.path_to_pptx}\n')
        for result in self.results_zip:
            if result[0] == 'Количество слайдов':
                txt_file.write(f'{result[0]} => {result[1]}\n')
            else:
                txt_file.write(f'{result[0]} => {"Да" if result[1] else "Нет"}\n')
        txt_file.close()
        return 'Успешная запись в файл'

    def csv(self, mode):
        """
        :param mode: Режим записи в файл
        :return: Возвращает str в случае успешного выполнения, иначе генерирует подробные ошибки
        """
        self._extension_check('.csv')
        mode = self._write_mode_check(mode)
        data_without_columns = [self.results_values]
        data_with_columns = [self.results]
        if self._empty(self.path_to_output, '.csv') or (not(self._empty(self.path_to_output, '.csv')) and mode == 'w'):
            self._write_to_csv(data_with_columns, self.path_to_output, mode, self.encoding)
            return 'Успешная запись в файл'
        elif not (self._empty(self.path_to_output, '.csv')) and mode == 'a':
            self._write_to_csv(data_without_columns, self.path_to_output, mode, self.encoding)
            return 'Успешная запись в файл'

    def excel(self, mode):
        pass

