# -*- coding: utf-8 -*-
import sys
import form

from CheckPresentationMain import CheckPresentationGetData as CPGD
from CheckPresentationMain import CheckPresentationAnalyze as CPA
from PyQt5.QtWidgets import QMainWindow, QApplication
from PyQt5.QtCore import QFileInfo
from PyQt5.QtGui import QPixmap
from pptx import Presentation


class CheckPresentationAppWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = form.Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.get_answer.clicked.connect(self.check_and_return_result)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):  # Catch drop event
        self.file = event.mimeData().urls()[0].toLocalFile()
        file_extension = QFileInfo(self.file).suffix()
        try:
            if file_extension == 'jpg' or file_extension == 'jpeg' or file_extension == 'png':
                self.ui.image_holder.setPixmap(QPixmap(self.file))
                self.ui.statusbar.append(f'Изображение по пути {self.file} установлено')
            elif file_extension == 'pptx':
                GetData = CPGD(Presentation(self.file))
                screens = GetData.generate_slide_images()
                top_words = GetData.analyze_text()
                self.ui.slide2_image_label.setPixmap(QPixmap(screens[2]))
                self.ui.slide3_image_label.setPixmap(QPixmap(screens[3]))
                self.ui.statusbar.append(f'Разбор pptx файла по пути {self.file}')
                self.ui.statusbar.append(f'Наиболее часто встречающиеся слова в презентации:')
                for word in top_words:
                    self.ui.statusbar.append(f'Слово "{word[0]}" встречается {word[1]} раз')
            else:
                self.ui.statusbar.append('Не поддерживаемое расширение файла.')
        except Exception as e:
            self.ui.statusbar.append(f'Ошибка {e} при загрузке файла.')

    def check_and_return_result(self):
        Analyze = CPA(Presentation(self.file))
        results = Analyze.analyze_results(
            txt_img_collisions_btn=self.ui.txt_img_collisions_btn_yes.isChecked(),
            distorted_images_btn=self.ui.distorted_images_btn_yes.isChecked(),
            all_collisions_btn=self.ui.all_collisions_btn_yes.isChecked(),
            content_compliance=self.ui.content_compliance_btn_yes.isChecked()
        )
        if self.ui.unloading_btn3.isChecked():
            for res in results:
                if res == 'Количество слайдов':
                    self.ui.statusbar.append(f'{res} => {results[res]}')
                else:
                    self.ui.statusbar.append(f'{res} => {"Да" if results[res] else "Нет"}')


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = CheckPresentationAppWindow()
    ex.show()
    sys.exit(app.exec_())
