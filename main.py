# -*- coding: utf-8 -*-
import sys
import form
from pptxChecker import PresentationUtils, PresentationCustomUtils
from PyQt5.QtWidgets import QMainWindow, QApplication
from PyQt5.QtCore import QFileInfo
from PyQt5.QtGui import QPixmap
from pptx import Presentation


class PresentationChecker(PresentationUtils):
    def get_slides_contents(self):
        slides = {
            1: {
                'textCounter': 0,
                'pictureCounter': 0,
                'titleCounter': 0,
            },
            2: {
                'textCounter': 0,
                'pictureCounter': 0,
                'titleCounter': 0,
            },
            3: {
                'textCounter': 0,
                'pictureCounter': 0,
                'titleCounter': 0,
            }
        }
        for slide in self.presentation.slides:
            index = int(self.presentation.slides.index(slide) + 1)
            if index > 3:
                return slides
            for shape in slide.shapes:
                if self.is_text(shape):
                    if self.is_title(shape):
                        slides[index]['titleCounter'] += 1
                    else:
                        slides[index]['textCounter'] += 1
                if self.is_image(shape):
                    slides[index]['pictureCounter'] += 1
        return slides

    def analyze_results(self, txt_img_collisions_btn=False, distorted_images_btn=False,  all_collisions_btn=False,
                              content_compliance=False):
        analyze_params = {
            'slides_count': self.get_slides_len(),
            'text_blocks_exist': None,
            'title_on_cover_page': None,
            'title_on_other_slides': None,
            'content_compliance': content_compliance,
            'single_typeface': None,
            'right_font_size': None,
            'text_not_overlaps_images': txt_img_collisions_btn,
            'images_not_distorted': distorted_images_btn,
            'images_not_overlaps_shapes': all_collisions_btn,
        }
        slides_contents = self.get_slides_contents()
        font_sizes = self.get_font_sizes()
        typefaces = self.get_typefaces()

        # Check first slide #
        slide1_font_size = 0
        slide1_blocks_correct = 0
        if max(font_sizes[1]) == 40 and min(font_sizes[1]) == 24:
            slide1_font_size = 1
        if slides_contents[1]['titleCounter'] + slides_contents[1]['textCounter'] == 2:
            slide1_blocks_correct = 1
        if slides_contents[1]['titleCounter'] >= 1 or \
                (slides_contents[1]['textCounter'] >= 1 and not slides_contents[1]['titleCounter']):
            analyze_params['title_on_cover_page'] = True
        # End first slide #

        # Check second slide #
        slide2_font_size = 0
        slide2_blocks_correct = 0
        slide2_title = 0
        if max(font_sizes[2]) == 24 and min(font_sizes[2]) == 20:
            slide2_font_size = 1
        if slides_contents[2]['textCounter'] + slides_contents[2]['titleCounter'] == 3 and \
                slides_contents[2]['pictureCounter'] == 2:
            slide2_blocks_correct = 1
            slide2_title = 1
        if slides_contents[2]['textCounter'] + slides_contents[2]['titleCounter'] == 2 and \
                slides_contents[2]['pictureCounter'] == 2:
            slide2_blocks_correct = 1
        # End second slide #

        # Check third slide #
        slide3_font_size = 0
        slide3_blocks_correct = 0
        slide3_title = 0
        if max(font_sizes[3]) == 24 and min(font_sizes[3]) == 20:
            slide3_font_size = 1
        if slides_contents[3]['textCounter'] + slides_contents[3]['titleCounter'] == 3 and \
                slides_contents[3]['pictureCounter'] == 3:
            slide3_blocks_correct = 1
        if slides_contents[3]['textCounter'] + slides_contents[3]['titleCounter'] == 4 and \
                slides_contents[3]['pictureCounter'] == 3:
            slide3_blocks_correct = 1
            slide3_title = 1
        if slides_contents[3]['titleCounter'] == 1:
            slide3_title = 1
        # End third slide #

        if (slide1_blocks_correct + slide2_blocks_correct + slide3_blocks_correct) == 3:
            analyze_params['text_blocks_exist'] = True
        else:
            analyze_params['text_blocks_exist'] = False
        if (slide2_title + slide3_title) == 2:
            analyze_params['title_on_other_slides'] = True
        else:
            analyze_params['title_on_other_slides'] = False
        if (slide1_font_size + slide2_font_size + slide3_font_size) == 3:
            analyze_params['right_font_size'] = True
        else:
            analyze_params['right_font_size'] = False
        if not len(typefaces) > 1:
            analyze_params['single_typeface'] = True
        else:
            analyze_params['single_typeface'] = False
        return analyze_params


class PCheckerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.Utils = PresentationCustomUtils()
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
                Check = PresentationChecker(Presentation(self.file))
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
        Check = PresentationChecker(Presentation(self.file))
        img_coll_btn = self.ui.txt_img_collisions_btn_yes.isChecked()
        dist_images = self.ui.distorted_images_btn_yes.isChecked()
        all_coll = self.ui.all_collisions_btn_yes.isChecked()
        cont_compliance = self.ui.content_compliance_btn_yes.isChecked()
        unloading_status_bar = self.ui.unloading_btn3.isChecked()
        results = Check.analyze_results(img_coll_btn, dist_images, all_coll, cont_compliance)
        translated_result = self.Utils.translate_results(results)
        if unloading_status_bar:
            for result in translated_result:
                self.ui.statusbar.append(f'{result} === {translated_result[result]}')


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = PCheckerWindow()
    ex.show()
    sys.exit(app.exec_())
