import sys
import form
import pygame
import base64
import os
import shutil
import PIL

from PIL import Image
from PyQt5.Qt import QMainWindow, QApplication, QFileInfo, QPixmap
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER_TYPE


class PCheckerUtils:
    def __init__(self):
        super().__init__()
        self.mso_pic = MSO_SHAPE_TYPE.PICTURE
        self.placeholder_pic = PP_PLACEHOLDER_TYPE.PICTURE

    @staticmethod
    def emu_px(emu):
        return round(emu // 9525)

    def get_width_height(self, presentation):
        return self.emu_px(presentation.slide_width), self.emu_px(presentation.slide_height)

    def save_images(self, presentation):
        slide_counter, image_cords, image_paths = 0, [], []
        try:  # creates dir, if exist delete and recreate #
            os.mkdir('img')
        except OSError:
            shutil.rmtree('img', ignore_errors=True)
            os.mkdir('img')
        for slide in presentation.slides:
            slide_counter += 1
            picture_counter = 1
            for shape in slide.shapes:
                if shape.shape_type == self.mso_pic or (shape.is_placeholder and shape.placeholder_format.type ==
                                                                                 self.placeholder_pic):
                    if slide_counter > 3:
                        raise Exception('Количество слайдов более чем 3, остановка парсинга картинок')
                    pil_pic_path = f"img/img_slide{slide_counter}_{picture_counter}.png"
                    pic_path     = f"img/img_slide{slide_counter}_{picture_counter}_original.png"
                    picture_counter += 1
                    with open(pic_path, 'wb') as f:
                        f.write(base64.b64decode(base64.b64encode(shape.image.blob)))
                        f.close()
                    pic_size = (self.emu_px(shape.width), self.emu_px(shape.height))
                    Image.open(pic_path).resize(pic_size, Image.ANTIALIAS).save(pil_pic_path)
                    os.remove(pic_path)
                    if 1 < slide_counter < 4:
                        image_cords.append((self.emu_px(shape.left), self.emu_px(shape.top)))
                        image_paths.append(pil_pic_path)

                else:
                    pass
        return [[image_paths[i], image_cords[i]] for i in range(len(image_cords))]

    def get_cords_dim(self, presentation):
        slide_counter = 0
        cords_dim = {'2': [], '3': []}
        for slide in presentation.slides:
            slide_counter += 1
            for shape in slide.shapes:
                if 1 < slide_counter < 4:
                    if hasattr(shape, "text"):
                        left_top     = (self.emu_px(shape.left), self.emu_px(shape.top))
                        width_height = (self.emu_px(shape.width), self.emu_px(shape.height))
                        cords_dim[str(slide_counter)].append([left_top, width_height])
        return cords_dim

    @staticmethod
    def create_screenshots(window_size, image_path_cords, prs_cords_dim_text):
        try:
            os.mkdir('screens')
        except OSError:
            shutil.rmtree('screens')
            os.mkdir('screens')
        slides_pic = {'2': [path for path in image_path_cords if 'slide2' in path[0]],
                      '3': [path for path in image_path_cords if 'slide3' in path[0]]}
        slides_nums = (2, 3)
        for slide in slides_pic:
            pygame.init()
            screen = pygame.display.set_mode(window_size)
            for text_cords in prs_cords_dim_text[slide]:
                pygame.draw.rect(screen, (255, 255, 255), (text_cords[0][0], text_cords[0][1],
                                                           text_cords[1][0], text_cords[1][1]), 1)
            for path_cords in slides_pic[slide]:
                screen.blit(pygame.image.load(path_cords[0]), path_cords[1])
            pygame.display.update()
            pygame.image.save(screen, f"screens/screen{slide}.jpg")
            pygame.quit()
        return f"screens/screen{slides_nums[0]}.jpg", f"screens/screen{slides_nums[1]}.jpg"


class PChecker(QMainWindow):
    def __init__(self):
        super().__init__()
        self.Utils = PCheckerUtils()
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
                slide_size         = self.Utils.get_width_height(Presentation(file))
                images_path_cords  = self.Utils.save_images(Presentation(file))
                prs_cords_dim_text = self.Utils.get_cords_dim(Presentation(file))
                screens            = self.Utils.create_screenshots(slide_size, images_path_cords, prs_cords_dim_text)
                self.ui.slide2_image_label.setPixmap(QPixmap(screens[0]))
                self.ui.slide3_image_label.setPixmap(QPixmap(screens[1]))

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
