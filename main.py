# -*- coding: utf-8 -*
import sys

from PyQt5 import uic
from PyQt5.Qt import QMainWindow, QApplication
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER_TYPE, MSO_SHAPE_TYPE

file_log = open('log.txt', 'w', encoding="utf-8")


def get_file_content(files):
    for file in files:
        slide_counter = 0  # Счётчик слайдов
        presentation = Presentation(f'{file}')
        presentation_slides_count = len(presentation.slides)  # Получаем общее количество слайдов

        res = {'Первый слайд': [], 'Второй слайд': [], 'Третий слайд': [], 'Другие слайды': []}
        file_log.write(f'Файл: {file}\n')
        for slide in presentation.slides:
            slide_counter += 1
            file_log.write(f'Начало {slide_counter} слайда:\n')
            for shape in slide.shapes:
                if slide_counter == 1:
                    if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
                        if shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.CENTER_TITLE:
                            res['Первый слайд'].append('Найден главный заголовок')
                            file_log.write(f'Найден главный заголовок:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.SUBTITLE:
                            res['Первый слайд'].append('Найден подзаголовок')
                            file_log.write(f'Найден подзаголовок:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.TITLE:
                            res['Первый слайд'].append('Найден заголовок')
                            file_log.write(f'Найден заголовок:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.BODY:
                            res['Первый слайд'].append('Найден body')
                            file_log.write(f'Найден body:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.OBJECT:
                            res['Первый слайд'].append('Найден обьект')
                            file_log.write(f'Найден обьект:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.PICTURE:
                            res['Первый слайд'].append('Найдена картинка')
                            file_log.write(f'Найдена картинка.\n')
                    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        res['Первый слайд'].append('Найдена картинка')
                        file_log.write(f'Найдена картинка.\n')
                    elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                        res['Первый слайд'].append('Найден auto_shape')
                        file_log.write(f'Найден auto_shape:\n')
                        if hasattr(shape, "text"):
                            file_log.write(f'{shape.text}\n')
                    elif shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                        res['Первый слайд'].append('Найден text_box')
                        file_log.write(f'Найден text_box:\n')
                        if hasattr(shape, "text"):
                            file_log.write(f'{shape.text}\n')
                    else:
                        print('Uncaught error')
                        break
                elif slide_counter == 2:
                    if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
                        if shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.CENTER_TITLE:
                            res['Второй слайд'].append('Найден главный заголовок')
                            file_log.write(f'Найден главный заголовок:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.SUBTITLE:
                            res['Второй слайд'].append('Найден подзаголовок')
                            file_log.write(f'Найден подзаголовок:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.TITLE:
                            res['Второй слайд'].append('Найден заголовок')
                            file_log.write(f'Найден заголовок:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.BODY:
                            res['Второй слайд'].append('Найден body')
                            file_log.write(f'Найден body:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.OBJECT:
                            res['Второй слайд'].append('Найден обьект')
                            file_log.write(f'Найден обьект:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.PICTURE:
                            res['Второй слайд'].append('Найдена картинка')
                            file_log.write(f'Найдена картинка.\n')
                    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        res['Второй слайд'].append('Найдена картинка')
                        file_log.write(f'Найдена картинка.\n')
                    elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                        res['Второй слайд'].append('Найден auto_shape')
                        file_log.write(f'Найден auto_shape:\n')
                        if hasattr(shape, "text"):
                            file_log.write(f'{shape.text}\n')
                    elif shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                        res['Второй слайд'].append('Найден text_box')
                        file_log.write(f'Найден text_box:\n')
                        if hasattr(shape, "text"):
                            file_log.write(f'{shape.text}\n')
                    else:
                        print('Uncaught error')
                        break
                elif slide_counter == 3:
                    if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
                        if shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.CENTER_TITLE:
                            res['Третий слайд'].append('Найден главный заголовок')
                            file_log.write(f'Найден главный заголовок:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.SUBTITLE:
                            res['Третий слайд'].append('Найден подзаголовок')
                            file_log.write(f'Найден подзаголовок:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.TITLE:
                            res['Третий слайд'].append('Найден заголовок')
                            file_log.write(f'Найден заголовок:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.BODY:
                            res['Третий слайд'].append('Найден body')
                            file_log.write(f'Найден body:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.OBJECT:
                            res['Третий слайд'].append('Найден обьект')
                            file_log.write(f'Найден обьект:\n')
                            if hasattr(shape, "text"):
                                file_log.write(f'{shape.text}\n')
                        elif shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.PICTURE:
                            res['Третий слайд'].append('Найдена картинка')
                            file_log.write(f'Найдена картинка.\n')
                    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        res['Третий слайд'].append('Найдена картинка')
                        file_log.write(f'Найдена картинка.\n')
                    elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                        res['Третий слайд'].append('Найден auto_shape')
                        file_log.write(f'Найден auto_shape:\n')
                        if hasattr(shape, "text"):
                            file_log.write(f'{shape.text}\n')
                    elif shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                        res['Третий слайд'].append('Найден text_box')
                        file_log.write(f'Найден text_box:\n')
                        if hasattr(shape, "text"):
                            file_log.write(f'{shape.text}\n')
                else:
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
