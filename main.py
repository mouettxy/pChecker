import sys
import form
import pygame
import base64
import os
import shutil
import PIL

from PIL import Image
from PyQt5.Qt import QMainWindow, QApplication, QWidget, QFileInfo, QPixmap, QPicture
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER_TYPE


def convert_emu_to_px(emu):
    return round(emu / 9525)


def get_prs_width_height(presentation):
    return convert_emu_to_px(presentation.slide_width), convert_emu_to_px(presentation.slide_height)


def get_and_save_images(presentation):
    slide_counter = 0
    image_cords = []
    image_paths = []
    # creates dir, if exist delete and recreate #
    try:
        os.mkdir('img')
    except:
        shutil.rmtree('img', ignore_errors=True)
        os.mkdir('img')
    for slide in presentation.slides:
        slide_counter += 1
        picture_counter = 1
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE or \
                    (shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER_TYPE.PICTURE):
                if slide_counter > 3:
                    raise Exception('Количество слайдов более чем 3, остановка парсинга картинок')
                pil_pic_path = f"img/img_slide{slide_counter}_{picture_counter}.png"
                pic_path = f"img/img_slide{slide_counter}_{picture_counter}_original.png"
                picture_counter += 1
                with open(pic_path, 'wb') as f:
                    f.write(base64.b64decode(base64.b64encode(shape.image.blob)))
                    f.close()
                pic_size = (convert_emu_to_px(shape.width), convert_emu_to_px(shape.height))
                img = Image.open(pic_path)
                img2 = img.resize(pic_size, Image.ANTIALIAS)
                img2.save(pil_pic_path)
                os.remove(pic_path)  # remove big files
                if 1 < slide_counter < 4:
                    image_cords.append((convert_emu_to_px(shape.left), convert_emu_to_px(shape.top)))
                    image_paths.append(pil_pic_path)

            else:
                pass
    matched_path_cords = [[image_paths[i], image_cords[i]] for i in range(len(image_cords))]
    return matched_path_cords


def get_cords_dim_text(presentation):
    slide_counter = 0
    prs_cords_dim_text = {'2': [], '3': []}
    for slide in presentation.slides:
        slide_counter += 1
        for shape in slide.shapes:
            if 1 < slide_counter < 4:
                if hasattr(shape, "text"):
                    left_top = (convert_emu_to_px(shape.left), convert_emu_to_px(shape.top))
                    width_height = (convert_emu_to_px(shape.width), convert_emu_to_px(shape.height))
                    text = shape.text
                    prs_cords_dim_text[str(slide_counter)].append([left_top, width_height, text])
    return prs_cords_dim_text

""" ############################################################################################################# FIX ME
def blit_text(surface, text, pos, font, color=pygame.Color('black')):
    words = [word.split(' ') for word in text.splitlines()]  # 2D array where each row is a list of words.
    space = font.size(' ')[0]  # The width of a space.
    max_width, max_height = surface.get_size()
    x, y = pos
    for line in words:
        for word in line:
            word_surface = font.render(word, 0, color)
            word_width, word_height = word_surface.get_size()
            if x + word_width >= max_width:
                x = pos[0]  # Reset the x.
                y += word_height  # Start on new row.
            surface.blit(word_surface, (x, y))
            x += word_width + space
        x = pos[0]  # Reset the x.
        y += word_height  # Start on new row.
"""

def create_screenshots(window_size, image_path_cords, prs_cords_dim_text):
    try:
        os.mkdir('screens')
    except:
        shutil.rmtree('screens', ignore_errors=True)
        os.mkdir('screens')
    slides_pic = {'2': [path for path in image_path_cords if 'slide2' in path[0]],
                  '3': [path for path in image_path_cords if 'slide3' in path[0]]}
    slides_nums = (2, 3)
    for slide in slides_pic:
        pygame.init()
        screen = pygame.display.set_mode(window_size)
        for path_cords in slides_pic[slide]:
            screen.blit(pygame.image.load(path_cords[0]), path_cords[1])
        pygame.display.update()
        pygame.image.save(screen, f"screens/screen{slide}.jpg")
        pygame.quit()
    return f"screens/screen{slides_nums[0]}", f"screens/screen{slides_nums[1]}"


class PChecker(QMainWindow):
    def __init__(self, parent=None):
        QWidget.__init__(self, parent)
        self.ui = form.Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.refresh_images.clicked.connect(self.update_images)
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
                images_path_cords = get_and_save_images(Presentation(file))
                prs_cords_dim_text = get_cords_dim_text(Presentation(file))
                screens = create_screenshots(slide_size, images_path_cords, prs_cords_dim_text)
                slide2_pixmap = QPixmap(screens[0])
                slide3_pixmap = QPixmap(screens[1])
                self.ui.slide2_image_label.setPixmap(slide2_pixmap)
                self.ui.slide3_image_label.setPixmap(slide3_pixmap)
                QApplication.processEvents()
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
