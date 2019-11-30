# -*- coding: utf-8 -*
import sys

from PyQt5 import uic
from PyQt5.Qt import QMainWindow, QApplication
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER_TYPE

file_log = open('log.txt', 'w', encoding="utf-8")


def get_file_content(files):
    for file in files:
        slide_counter = 0  # Счётчик слайдов
        presentation = Presentation(f'{file}')
        presentation_slides_number = len(presentation.slides)  # Получаем общее количество слайдов

        presentation_slide_objects = {'1': [], '2': [], '3': [], '4': []}
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
        possible_shapes_placeholders = [
            "PP_PLACEHOLDER_TYPE.CENTER_TITLE",
            "PP_PLACEHOLDER_TYPE.SUBTITLE",
            "PP_PLACEHOLDER_TYPE.TITLE",
            "PP_PLACEHOLDER_TYPE.BODY",
            "PP_PLACEHOLDER_TYPE.OBJECT",
            "PP_PLACEHOLDER_TYPE.PICTURE",
        ]
        possible_shapes_mso = [
            "MSO_SHAPE_TYPE.PICTURE",
            "MSO_SHAPE_TYPE.AUTO_SHAPE",
            "MSO_SHAPE_TYPE.TEXT_BOX",
        ]
        finded_text = []
        more_than_need_slides = False
        # file_log.write(f'Файл: {file}\n')
        for slide in presentation.slides:
            slide_counter += 1
            # file_log.write(f'Начало {slide_counter} слайда:\n')
            for shape in slide.shapes:
                for number in range(1, presentation_slides_number + 1):
                    if slide_counter == number and number < 4:
                        if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
                            for possible_shape_placeholder in possible_shapes_placeholders:
                                if shape.placeholder_format.type is eval(possible_shape_placeholder):
                                    if hasattr(shape, "text"):
                                        finded_text.append(shape.text)
                                        if hasattr(shape, "text_frame"):
                                            font_sizes = []
                                            if font_sizes:
                                                font_sizes.clear()
                                            for paragraph in shape.text_frame.paragraphs:
                                                for run in paragraph.runs:
                                                    try:
                                                        font_sizes.append(run.font.size.pt)
                                                    except AttributeError:
                                                        pass
                                            else:
                                                if font_sizes:
                                                    presentation_slide_objects[f'{number}'].append(
                                                        {'Shape': possible_shape_placeholder,
                                                         'HasText': hasattr(shape, "text"),
                                                         'TextSize': min(font_sizes),
                                                         })
                                                else:
                                                    presentation_slide_objects[f'{number}'].append(
                                                        {'Shape': possible_shape_placeholder,
                                                         'HasText': hasattr(shape, "text"),
                                                         'TextSize': None,
                                                         })
                                    else:
                                        presentation_slide_objects[f'{number}'].append(
                                            {'Shape': possible_shape_placeholder,
                                             'HasText': None,
                                             'TextSize': None,
                                             })
                        else:
                            for possible_shape_mso in possible_shapes_mso:
                                if shape.shape_type is eval(possible_shape_mso):
                                    if hasattr(shape, "text"):
                                        if hasattr(shape, "text_frame"):
                                            finded_text.append(shape.text)
                                            font_sizes = []
                                            if font_sizes:
                                                font_sizes.clear()
                                            for paragraph in shape.text_frame.paragraphs:
                                                for run in paragraph.runs:
                                                    try:
                                                        font_sizes.append(run.font.size.pt)
                                                    except AttributeError:
                                                        pass
                                            else:
                                                if font_sizes:
                                                    presentation_slide_objects[f'{number}'].append(
                                                        {'Shape': possible_shape_mso,
                                                         'HasText': hasattr(shape, "text"),
                                                         'TextSize': min(font_sizes),
                                                         })
                                                else:
                                                    presentation_slide_objects[f'{number}'].append(
                                                        {'Shape': possible_shape_mso,
                                                         'HasText': hasattr(shape, "text"),
                                                         'TextSize': None,
                                                         })
                                    else:
                                        presentation_slide_objects[f'{number}'].append({'Shape': possible_shape_mso,
                                                                                        'HasText': None,
                                                                                        'TextSize': None,
                                                                                        })
                    elif number >= 4:
                        more_than_need_slides = True

        text_sizes = {'1': [], '2': [], '3': []}
        title_exist = False
        text_counter = 0
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
                                text_sizes['1'].append(slide_shapes[slide_shape])
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
                                text_sizes['2'].append(slide_shapes[slide_shape])
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
                                text_sizes['3'].append(slide_shapes[slide_shape])
                    else:
                        more_than_need_slides = True

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
        get_file_content(files)
        file_log.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = PChecker()
    ex.show()
    sys.exit(app.exec_())
