import sys
import xlsxwriter
import re
import form


from nltk import FreqDist
from PyQt5 import uic
from PyQt5.Qt import QMainWindow, QApplication, QWidget, QFileInfo, QPixmap
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER_TYPE


def convert_emu_to_px(emu):
    return round(emu / 9525)


def get_prs_width_height(presentation):
    return convert_emu_to_px(presentation.slide_width), convert_emu_to_px(presentation.slide_height)


def get_images(presentation):
    slide_counter = 0
    image_blob_cords = {'2': (), '3': ()}
    for slide in presentation.slides:
        slide_counter += 1
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE or \
                    (shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.PICTURE):
                if slide_counter > 3:
                    raise Exception('Количество слайдов более чем 3, остановка парсинга картинок')
                picture_blob = shape.image.blob
                if len(picture_blob) == 0:
                    raise Exception('Не найдена картинка')
                top_left_cords = (convert_emu_to_px(shape.top), convert_emu_to_px(shape.left))
                if 1 < slide_counter < 4:
                    image_blob_cords[str(slide_counter)] = picture_blob, top_left_cords
            else:
                pass
    return image_blob_cords


def create_screenshots(window_size, blob_and_cords):
    pass


class PChecker(QMainWindow):
    def __init__(self, parent=None):
        QWidget.__init__(self, parent)
        self.ui = form.Ui_MainWindow()
        self.ui.setupUi(self)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):  # Ловим ивент дропа файла в окно
        file = event.mimeData().urls()[0].toLocalFile()
        file_extension = QFileInfo(file).suffix()
        try:
            if file_extension == 'jpg' or file_extension == 'jpeg' or file_extension == 'png':
                self.ui.image_holder.setPixmap(QPixmap(file))
                self.ui.statusbar.append(f'Изображение по пути {file} установлено')
            elif file_extension == 'pptx':
                slide_size = get_prs_width_height(Presentation(file))
                images_blob_cords = get_images(Presentation(file))
                create_screenshots(slide_size, images_blob_cords)
                self.ui.statusbar.append(f'Разбор pptx файла по пути {file}')
            else:
                self.ui.statusbar.append('Не поддерживаемое расширение файла.')
        except Exception as e:
            self.ui.statusbar.append(f'Ошибка {e} при загрузке файла.')


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = PChecker()
    ex.show()
    sys.exit(app.exec_())
