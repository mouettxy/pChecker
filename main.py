# -*- coding: utf-8 -*-
import sys
import form
from pChecker import PChecker, PCheckerUtils
from PyQt5.QtWidgets import QMainWindow, QApplication
from PyQt5.QtCore import QFileInfo
from PyQt5.QtGui import QPixmap
from pptx import Presentation


class PCheckerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.Utils = PCheckerUtils()
        self.ui = form.Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.get_answer.clicked.connect(self.check_and_return_result)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):  # Ловим ивент дропа файла в окно
        self.file = event.mimeData().urls()[0].toLocalFile()
        file_extension = QFileInfo(self.file).suffix()
        try:
            if file_extension == 'jpg' or file_extension == 'jpeg' or file_extension == 'png':
                self.ui.image_holder.setPixmap(QPixmap(self.file))
                self.ui.statusbar.append(f'Изображение по пути {self.file} установлено')
            elif file_extension == 'pptx':
                Check = PChecker(Presentation(self.file))
                slide_size         = self.Utils.get_width_height(Presentation(self.file))
                images_path_cords  = self.Utils.save_images(Presentation(self.file))
                prs_cords_dim_text = self.Utils.get_cords_dim(Presentation(self.file))
                screens            = self.Utils.create_screenshots(slide_size, images_path_cords, prs_cords_dim_text)
                content_compliance = Check.analyze_text()
                content_compliance = [words for words in content_compliance]
                self.ui.slide2_image_label.setPixmap(QPixmap(screens[0]))
                self.ui.slide3_image_label.setPixmap(QPixmap(screens[1]))
                self.ui.statusbar.append(f'Разбор pptx файла по пути {self.file}')
                self.ui.statusbar.append(f'Наиболее часто встречающиеся слова в презентации:')
                for tuple_content in content_compliance:
                    self.ui.statusbar.append(f'{tuple_content[0]} => {tuple_content[1]}')
            else:
                self.ui.statusbar.append('Не поддерживаемое расширение файла.')
        except Exception as e:
            self.ui.statusbar.append(f'Ошибка {e} при загрузке файла.')

    def check_and_return_result(self):
        Check = PChecker(Presentation(self.file))
        img_coll_btn = self.ui.txt_img_collisions_btn_yes.isChecked()
        dist_images = self.ui.distorted_images_btn_yes.isChecked()
        all_coll = self.ui.all_collisions_btn_yes.isChecked()
        cont_compliance = self.ui.content_compliance_btn_yes.isChecked()
        unloading_txt = self.ui.unloading_btn1.isChecked()
        unloading_xlsx = self.ui.unloading_btn2.isChecked()
        unloading_statusbar = self.ui.unloading_btn3.isChecked()
        results = Check.analyze_results(img_coll_btn, dist_images, all_coll, cont_compliance)
        translated_result = self.Utils.translate_results(results)
        if unloading_statusbar:
            for result in translated_result:
                self.ui.statusbar.append(f'{result} === {translated_result[result]}')
        elif unloading_txt:
            Check.unloading_txt()
        else:
            Check.unloading_xlsx()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = PCheckerWindow()
    ex.show()
    sys.exit(app.exec_())
