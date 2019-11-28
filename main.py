# -*- coding: utf-8 -*
import sys

from PyQt5 import uic
from PyQt5.Qt import QMainWindow, QApplication
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

file_log = open('log.txt', 'w', encoding="utf-8")


def get_file_content(files):
    for file in files:
        slide_counter = 0  # Счётчик слайдов
        presentation = Presentation(f'{file}')
        presentation_slides_number = len(presentation.slides)  # Получаем общее количество слайдов

        res = {'1 слайд': [], '2 слайд': [], '3 слайд': [], 'Другие слайды': []}
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
        file_log.write(f'Файл: {file}\n')
        for slide in presentation.slides:
            slide_counter += 1
            file_log.write(f'Начало {slide_counter} слайда:\n')
            for shape in slide.shapes:
                for number in range(1, presentation_slides_number + 1):
                    if slide_counter == number and number < 4:
                        if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
                            for possible_shape_placeholder in possible_shapes_placeholders:
                                if shape.placeholder_format.type is eval(possible_shape_placeholder):
                                    res[f'{number} слайд'].append(f' Найден объект {possible_shape_placeholder} \n')
                                    file_log.write(f' Найден обьект {possible_shape_placeholder} \n')
                                    if hasattr(shape, "text"):
                                        file_log.write(f'  Текст на объекте: {shape.text} \n')
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
                                                    file_log.write(f'   Размер текста {max(font_sizes)} \n')
                        else:
                            for possible_shape_mso in possible_shapes_mso:
                                if shape.shape_type is eval(possible_shape_mso):
                                    res[f'{number} слайд'].append(f' Найден объект {possible_shape_mso} \n')
                                    file_log.write(f' Найден обьект {possible_shape_mso} \n')
                                    if hasattr(shape, "text"):
                                        file_log.write(f'  Текст на объекте: {shape.text} \n')
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
                                                    file_log.write(f'   Размер текста {max(font_sizes)} \n')
                    elif number >= 4:
                        res['Другие слайды'].append('Найдены.')
                        file_log.write(f'Найдены слайды больше 4го.\n')

        for result in res:
            print(f'{result} -> {res.get(result)}')


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
