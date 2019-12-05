# -*- coding: utf-8 -*
import sys
import xlsxwriter

from PyQt5 import uic
from PyQt5.Qt import QMainWindow, QApplication
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER_TYPE

workbook = xlsxwriter.Workbook('result.xlsx')
worksheet = workbook.add_worksheet()


def get_file_content(files, theme):
    row = 0
    col = 0
    iterator_counter = -1
    for file in files:
        print(file)
        iterator_counter += 1
        print(theme)
        presentation = Presentation(f'{file}')
        detailed_total_score = {
            'Слайдов': None,
            'Блоки текста и изображений размещены': None,
            'Название на титульном': None,
            'Заголовки на 2 и 3 слайдах': None,
            'Соответствие содержанию': None,
            'Единый шрифт': None,
            'Верный размер шрифта': None,
            'Текст не перекрывает изображения': None,
            'Изображения не искажены': None,
            'Изображения не перекрывают текст, заголовок, друг друга': None
        }
        possible_shapes = {
            'PLACEHOLDERS': [
                "PP_PLACEHOLDER_TYPE.CENTER_TITLE",
                "PP_PLACEHOLDER_TYPE.SUBTITLE",
                "PP_PLACEHOLDER_TYPE.TITLE",
                "PP_PLACEHOLDER_TYPE.BODY",
                "PP_PLACEHOLDER_TYPE.OBJECT",
                "PP_PLACEHOLDER_TYPE.PICTURE",
            ],
            'MSO': [
                "MSO_SHAPE_TYPE.PICTURE",
                "MSO_SHAPE_TYPE.AUTO_SHAPE",
                "MSO_SHAPE_TYPE.TEXT_BOX",
            ]
        }
        presentation_slide_objects = {'1': [], '2': [], '3': [], '4': []}
        slide_counter = 0
        finded_text = set()
        text_typeface = set()
        for slide in presentation.slides:
            slide_counter += 1
            for shape in slide.shapes:
                for shape_type in possible_shapes:
                    for type_ in possible_shapes[shape_type]:
                        if not slide_counter >= 3:
                            if shape.is_placeholder:
                                if hasattr(shape, "text_frame"):
                                    finded_text.add(shape.text)
                                    font_sizes = []
                                    if font_sizes:
                                        font_sizes.clear()
                                    for paragraph in shape.text_frame.paragraphs:
                                        for run in paragraph.runs:
                                            try:
                                                font_sizes.append(run.font.size.pt)
                                                text_typeface.add(run.font.name)
                                            except AttributeError:
                                                pass
                                        else:
                                            if font_sizes:
                                                presentation_slide_objects[f'{slide_counter}'].append(
                                                    {'Shape': type_,
                                                     'HasText': hasattr(shape, "text"),
                                                     'TextSize': min(font_sizes),
                                                     })
                                            else:
                                                presentation_slide_objects[f'{slide_counter}'].append(
                                                    {'Shape': type_,
                                                     'HasText': hasattr(shape, "text"),
                                                     'TextSize': None,
                                                     })

                                    else:
                                        presentation_slide_objects[f'{slide_counter}'].append(
                                            {'Shape': type_,
                                             'HasText': None,
                                             'TextSize': None,
                                             })
                                else:
                                    if hasattr(shape, "text_frame"):
                                        finded_text.add(shape.text)
                                        font_sizes = []
                                        if font_sizes:
                                            font_sizes.clear()
                                        for paragraph in shape.text_frame.paragraphs:
                                            for run in paragraph.runs:
                                                try:
                                                    font_sizes.append(run.font.size.pt)
                                                    text_typeface.add(run.font.name)
                                                except AttributeError:
                                                    pass
                                        else:
                                            if font_sizes:
                                                presentation_slide_objects[f'{slide_counter}'].append(
                                                    {'Shape': type_,
                                                     'HasText': hasattr(shape, "text"),
                                                     'TextSize': min(font_sizes),
                                                     })
                                            else:
                                                presentation_slide_objects[f'{slide_counter}'].append(
                                                    {'Shape': type_,
                                                     'HasText': hasattr(shape, "text"),
                                                     'TextSize': None,
                                                     })
                                    else:
                                        presentation_slide_objects[f'{slide_counter}'].append({'Shape': type_,
                                                                                               'HasText': None,
                                                                                               'TextSize': None,
                                                                                               })
                            else:
                                pass
        text_sizes = {'1': set(), '2': set(), '3': set()}
        title_exist = False
        text_counter = 0
        text_size_counter = False
        text_counter_only_2_3_slides = 0
        image_counter = 0
        for slide_score in presentation_slide_objects:
            for slide_shapes in presentation_slide_objects[slide_score]:
                for slide_shape in slide_shapes:
                    if slide_score == '1':
                        if slide_shape == 'HasText':
                            if slide_shapes[slide_shape]:
                                text_counter += 1
                        if slide_shape == 'TextSize':
                            if slide_shapes[slide_shape]:
                                if slide_shapes[slide_shape] >= 30:
                                    title_exist = True
                                text_sizes['1'].add(slide_shapes[slide_shape])
                    elif slide_score == '2':
                        if slide_shape == 'HasText':
                            if slide_shapes[slide_shape]:
                                text_counter += 1
                                text_counter_only_2_3_slides += 1
                        if slide_shapes[slide_shape] == 'PP_PLACEHOLDER_TYPE.PICTURE' or \
                                slide_shapes[slide_shape] == 'MSO_SHAPE_TYPE.PICTURE':
                            image_counter += 1
                        if slide_shape == 'TextSize':
                            if slide_shapes[slide_shape]:
                                text_sizes['2'].add(slide_shapes[slide_shape])
                    elif slide_score == '3':
                        if slide_shape == 'HasText':
                            if slide_shapes[slide_shape]:
                                text_counter += 1
                                text_counter_only_2_3_slides += 1
                        if slide_shapes[slide_shape] == 'PP_PLACEHOLDER_TYPE.PICTURE' or \
                                slide_shapes[slide_shape] == 'MSO_SHAPE_TYPE.PICTURE':
                            image_counter += 1
                        if slide_shape == 'TextSize':
                            if slide_shapes[slide_shape]:
                                text_sizes['3'].add(slide_shapes[slide_shape])
        for text in finded_text:
            if theme.lower() or theme in text:
                content_compliance = True
            else:
                content_compliance = False

        if (40 in text_sizes['1']) and (24 or 20 in text_sizes['2']) and (24 or 20 in text_sizes['3']):
            text_size_counter = True

        detailed_total_score['Слайдов'] = slide_counter
        if text_counter >= 7:
            detailed_total_score['Блоки текста и изображений размещены'] = 'Да'
        else:
            detailed_total_score['Блоки текста и изображений размещены'] = 'Нет'
        if title_exist:
            detailed_total_score['Название на титульном'] = 'Да'
        else:
            detailed_total_score['Название на титульном'] = 'Нет'
        if text_counter_only_2_3_slides == 7:
            detailed_total_score['Заголовки на 2 и 3 слайдах'] = 'Да'
        else:
            detailed_total_score['Заголовки на 2 и 3 слайдах'] = 'Нет'
        if content_compliance:
            detailed_total_score['Соответствие содержанию'] = 'Да'
        else:
            detailed_total_score['Соответствие содержанию'] = 'Нет'
        if len(text_typeface) == 1:
            detailed_total_score['Единый шрифт'] = 'Да'
        else:
            detailed_total_score['Единый шрифт'] = 'Нет'
        if text_size_counter:
            detailed_total_score['Верный размер шрифта'] = 'Да'
        else:
            detailed_total_score['Верный размер шрифта'] = 'Нет'

        if not iterator_counter:
            for detailed_total_score_title in detailed_total_score:
                worksheet.write(row, col, detailed_total_score_title)
                col += 1
            row += 1
        col = 0
        for score in detailed_total_score:
            worksheet.write(row, col, detailed_total_score[score])
            col += 1
        row += 1

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
        get_file_content(files, self.project_theme.text())
        workbook.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = PChecker()
    ex.show()
    sys.exit(app.exec_())
