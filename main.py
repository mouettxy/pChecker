# -*- coding: utf-8 -*
import sys
import xlsxwriter
import re

from nltk import FreqDist
from PyQt5 import uic
from PyQt5.Qt import QMainWindow, QApplication
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER_TYPE


def get_file_content(files):
    """
    :param files:
    :return: objects on 1, 2, 3 slides. Objects have object ID, boolean hasText, textSize(pt) \
    :return: is a dict that slide number is a key and objects is a list
    """
    for file in files:
        presentation = Presentation(f'{file}')
        presentation_slide_objects = {'1': [], '2': [], '3': []}
        slides = presentation.slides
        slides_len = len(presentation.slides)
        all_text = set()
        text_typefaces = set()
        for slide in slides:
            slide_index = slides.index(slide) + 1
            if slide_index != 4:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        all_text.add(shape.text.strip().lower())
                        font_sizes = []
                        if font_sizes:
                            font_sizes.clear()
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                try:
                                    font_sizes.append(run.font.size.pt)
                                    text_typefaces.add(run.font.name)
                                except AttributeError:
                                    continue
                        else:
                            if font_sizes:
                                presentation_slide_objects[f'{slide_index}'].append(
                                    {'Shape': shape.shape_type,
                                     'HasText': hasattr(shape, "text"),
                                     'TextSize': min(font_sizes),
                                     })
                            else:
                                presentation_slide_objects[f'{slide_index}'].append(
                                    {'Shape': shape.shape_type,
                                     'HasText': hasattr(shape, "text"),
                                     'TextSize': None,
                                     })
                    else:
                        presentation_slide_objects[f'{slide_index}'].append(
                            {'Shape': shape.shape_type,
                             'HasText': None,
                             'TextSize': None,
                             })
            else:
                continue
        else:
            return presentation_slide_objects, slides_len, all_text, text_typefaces


def get_slides_result(prs_info, presentation_theme):
    """
    :param prs_info:
    :param presentation_theme:
    :return: dict that have a number of slides, presentation theme, count of text, count of images
    """
    presentation_objects = prs_info[0]
    presentation_slide_length = prs_info[1]
    presentation_text_on_slides = prs_info[2]
    presentation_text_typefaces = prs_info[3]
    text_sizes = {'1': set(), '2': set(), '3': set()}
    text_counter_1 = 0
    text_counter_2 = 0
    image_counter_2 = 0
    text_counter_3 = 0
    image_counter_3 = 0
    title_exist = False
    for slide_index in presentation_objects:
        for slide_objects in presentation_objects[slide_index]:
            for slide_object in slide_objects:
                if slide_index == '1':
                    if slide_object == 'HasText':
                        if slide_objects[slide_object]:
                            text_counter_1 += 1
                    elif slide_object == 'TextSize':
                        if slide_objects[slide_object]:
                            if int(slide_objects[slide_object]) > 30:
                                title_exist = True
                            text_sizes['1'].add(slide_objects[slide_object])
                elif slide_index == '2':
                    if slide_object == 'HasText':
                        if slide_objects[slide_object]:
                            text_counter_2 += 1
                        else:
                            image_counter_2 += 1
                    elif slide_object == 'TextSize':
                        if slide_objects[slide_object]:
                            text_sizes['2'].add(slide_objects[slide_object])
                else:
                    if slide_object == 'HasText':
                        if slide_objects[slide_object]:
                            text_counter_3 += 1
                        else:
                            image_counter_3 += 1
                    elif slide_object == 'TextSize':
                        if slide_objects[slide_object]:
                            text_sizes['3'].add(slide_objects[slide_object])
    else:

        # Get content compliance #
        all_text = []
        for text in presentation_text_on_slides:
            text_formatted = re.sub(r"\d+", "", (re.sub(r"\b\w{0,3}\b", "", text.lower()).strip()))
            text_formatted = re.sub(r"[-.?!)(,:]", "", text_formatted)
            all_text.extend(text_formatted.split())
        top_3_words = FreqDist(all_text).most_common(3)
        for word in top_3_words:
            theme_agreed = True if presentation_theme.lower() in word[0] else False
            if theme_agreed:
                break
        # End get content compliance. Answer in bool theme_agreed #

        # Get text and images count #
        total_text_images_count = int(
            text_counter_1 + text_counter_2 + text_counter_3 + image_counter_2 + image_counter_3)

        if total_text_images_count == 12:
            text_and_images_exist = True
            title_2_3_exist = False
        elif total_text_images_count == 14:
            text_and_images_exist = True
            title_2_3_exist = True
        else:
            text_and_images_exist = False
            title_2_3_exist = False
        # End get text and images count. Answer in bool text_and_images_exist and title_2_3_exist #

        # Get font size #
        font_size = True if (40 in text_sizes['1']) and \
                            (24 or 20 in text_sizes['2']) and (24 or 20 in text_sizes['3']) else False
        # End get font size. Answer in bool font size #

        # Total #
        answer = {
            'Слайдов': presentation_slide_length,
            'Блоки текста и изображений размещены': 'Да' if text_and_images_exist else 'Нет',
            'Название на титульном': 'Да' if title_exist else 'Нет',
            'Заголовки на 2 и 3 слайдах': 'Да' if title_2_3_exist else 'Нет',
            'Соответствие содержанию': 'Да' if theme_agreed else 'Нет',
            'Единый шрифт': 'Да' if len(presentation_text_typefaces) == 1 else 'Нет',
            'Верный размер шрифта': 'Да' if font_size else 'Нет',
            'Текст не перекрывает изображения': None,
            'Изображения не искажены': None,
            'Изображения не перекрывают текст, заголовок, друг друга': None
        }
        # End Total #

        return answer


def add_to_excel(answer):
    workbook = xlsxwriter.Workbook('result.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0
    col = 1
    for key in answer:
        worksheet.write(row, col, key)
        worksheet.write(row + 1, col, answer[key])
        col += 1
    workbook.close()


class PChecker(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('form.ui', self)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):  # Ловим ивент дропа файла в окно
        files = [u.toLocalFile() for u in event.mimeData().urls()]  # Получаем пути файлов
        add_to_excel(get_slides_result(get_file_content(files), self.project_theme.text()))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = PChecker()
    ex.show()
    sys.exit(app.exec_())
