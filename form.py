# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'form.ui'
#
# Created by: PyQt5 UI code generator 5.13.0
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1020, 680)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setStyleSheet("QVBoxLayout {\n"
"                    border: none;\n"
"                    }\n"
"                ")
        self.centralwidget.setObjectName("centralwidget")
        self.main_tab_widget = QtWidgets.QTabWidget(self.centralwidget)
        self.main_tab_widget.setGeometry(QtCore.QRect(-1, 0, 1024, 700))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.main_tab_widget.sizePolicy().hasHeightForWidth())
        self.main_tab_widget.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        self.main_tab_widget.setFont(font)
        self.main_tab_widget.setStyleSheet("border: none;")
        self.main_tab_widget.setTabBarAutoHide(True)
        self.main_tab_widget.setObjectName("main_tab_widget")
        self.single_file_check = QtWidgets.QWidget()
        self.single_file_check.setStyleSheet("background-color: #263238;")
        self.single_file_check.setObjectName("single_file_check")
        self.input_data = QtWidgets.QTextEdit(self.single_file_check)
        self.input_data.setGeometry(QtCore.QRect(10, 10, 661, 101))
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(-1)
        self.input_data.setFont(font)
        self.input_data.setStyleSheet("color: #FAFAFA;\n"
"                                font-size: 12px;\n"
"                                background-color: #263238;\n"
"                                border: none;\n"
"                            ")
        self.input_data.setObjectName("input_data")
        self.slides_box = QtWidgets.QToolBox(self.single_file_check)
        self.slides_box.setGeometry(QtCore.QRect(10, 120, 671, 451))
        self.slides_box.setStyleSheet("border: 1px solid #FAFAFA;\n"
"                                color: #FAFAFA;\n"
"                                font-size: 12px;\n"
"                                font-weight: bold;\n"
"                                padding: 4px;\n"
"                                border-radius: 2px;\n"
"                            ")
        self.slides_box.setObjectName("slides_box")
        self.slide1_page = QtWidgets.QWidget()
        self.slide1_page.setGeometry(QtCore.QRect(0, 0, 651, 375))
        self.slide1_page.setStyleSheet("border: none;")
        self.slide1_page.setObjectName("slide1_page")
        self.slide2_image_label = QtWidgets.QLabel(self.slide1_page)
        self.slide2_image_label.setGeometry(QtCore.QRect(0, 0, 651, 371))
        self.slide2_image_label.setText("")
        self.slide2_image_label.setTextFormat(QtCore.Qt.PlainText)
        self.slide2_image_label.setScaledContents(True)
        self.slide2_image_label.setObjectName("slide2_image_label")
        self.slides_box.addItem(self.slide1_page, "")
        self.slide2_page = QtWidgets.QWidget()
        self.slide2_page.setGeometry(QtCore.QRect(0, 0, 651, 375))
        self.slide2_page.setStyleSheet("border: none;")
        self.slide2_page.setObjectName("slide2_page")
        self.slide3_image_label = QtWidgets.QLabel(self.slide2_page)
        self.slide3_image_label.setGeometry(QtCore.QRect(0, 0, 651, 381))
        self.slide3_image_label.setText("")
        self.slide3_image_label.setTextFormat(QtCore.Qt.PlainText)
        self.slide3_image_label.setScaledContents(True)
        self.slide3_image_label.setWordWrap(False)
        self.slide3_image_label.setObjectName("slide3_image_label")
        self.slides_box.addItem(self.slide2_page, "")
        self.statusbar = QtWidgets.QTextEdit(self.single_file_check)
        self.statusbar.setGeometry(QtCore.QRect(10, 580, 1001, 71))
        self.statusbar.setStyleSheet("border: none;\n"
"                                padding-left: 4px;\n"
"                                padding-right: 4px;\n"
"                                color: #212121;\n"
"                                background-color: #546E7A;\n"
"                                border-radius: 2px;\n"
"                            ")
        self.statusbar.setObjectName("statusbar")
        self.get_answer = QtWidgets.QPushButton(self.single_file_check)
        self.get_answer.setGeometry(QtCore.QRect(690, 540, 321, 31))
        self.get_answer.setStyleSheet("border-radius: 2px;\n"
"                                background-color: #455A64;\n"
"                                border: none;\n"
"                                color: #FAFAFA;\n"
"                                font-size: 12px;\n"
"                                font-weight: bold;\n"
"                            ")
        self.get_answer.setObjectName("get_answer")
        self.image_holder_label = QtWidgets.QLabel(self.single_file_check)
        self.image_holder_label.setGeometry(QtCore.QRect(690, 10, 251, 20))
        self.image_holder_label.setStyleSheet("border: none;\n"
"                                color: #FAFAFA;\n"
"                                font-weight: bold;\n"
"                                font-size: 12px;\n"
"                            ")
        self.image_holder_label.setObjectName("image_holder_label")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.single_file_check)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(690, 210, 160, 80))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.txt_img_collisions_layout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.txt_img_collisions_layout.setContentsMargins(0, 0, 0, 0)
        self.txt_img_collisions_layout.setObjectName("txt_img_collisions_layout")
        self.txt_img_collisions_label = QtWidgets.QLabel(self.verticalLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(-1)
        font.setBold(True)
        font.setWeight(75)
        self.txt_img_collisions_label.setFont(font)
        self.txt_img_collisions_label.setStyleSheet("border: none;\n"
"                                            color: #FAFAFA;\n"
"                                            font-weight: bold;\n"
"                                            font-size: 12px;\n"
"                                        ")
        self.txt_img_collisions_label.setWordWrap(True)
        self.txt_img_collisions_label.setObjectName("txt_img_collisions_label")
        self.txt_img_collisions_layout.addWidget(self.txt_img_collisions_label)
        self.txt_img_collisions_btn_yes = QtWidgets.QRadioButton(self.verticalLayoutWidget)
        self.txt_img_collisions_btn_yes.setStyleSheet("color: #FAFAFA;\n"
"                                            border: none;\n"
"                                            font-size: 11px;\n"
"                                        ")
        self.txt_img_collisions_btn_yes.setObjectName("txt_img_collisions_btn_yes")
        self.txt_img_collisions_layout.addWidget(self.txt_img_collisions_btn_yes)
        self.txt_img_collisions_btn_no = QtWidgets.QRadioButton(self.verticalLayoutWidget)
        self.txt_img_collisions_btn_no.setStyleSheet("color: #FAFAFA;\n"
"                                            border: none;\n"
"                                            font-size: 11px;\n"
"                                        ")
        self.txt_img_collisions_btn_no.setChecked(True)
        self.txt_img_collisions_btn_no.setObjectName("txt_img_collisions_btn_no")
        self.txt_img_collisions_layout.addWidget(self.txt_img_collisions_btn_no)
        self.verticalLayoutWidget_2 = QtWidgets.QWidget(self.single_file_check)
        self.verticalLayoutWidget_2.setGeometry(QtCore.QRect(860, 210, 160, 80))
        self.verticalLayoutWidget_2.setObjectName("verticalLayoutWidget_2")
        self.distorted_images_layout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget_2)
        self.distorted_images_layout.setContentsMargins(0, 0, 0, 0)
        self.distorted_images_layout.setObjectName("distorted_images_layout")
        self.distorted_images_label = QtWidgets.QLabel(self.verticalLayoutWidget_2)
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(-1)
        font.setBold(True)
        font.setWeight(75)
        self.distorted_images_label.setFont(font)
        self.distorted_images_label.setStyleSheet("border: none;\n"
"                                            color: #FAFAFA;\n"
"                                            font-weight: bold;\n"
"                                            font-size: 12px;\n"
"                                        ")
        self.distorted_images_label.setWordWrap(True)
        self.distorted_images_label.setObjectName("distorted_images_label")
        self.distorted_images_layout.addWidget(self.distorted_images_label)
        self.distorted_images_btn_yes = QtWidgets.QRadioButton(self.verticalLayoutWidget_2)
        self.distorted_images_btn_yes.setStyleSheet("color: #FAFAFA;\n"
"                                            border: none;\n"
"                                            font-size: 11px;\n"
"                                        ")
        self.distorted_images_btn_yes.setObjectName("distorted_images_btn_yes")
        self.distorted_images_layout.addWidget(self.distorted_images_btn_yes)
        self.distorted_images_btn_no = QtWidgets.QRadioButton(self.verticalLayoutWidget_2)
        self.distorted_images_btn_no.setStyleSheet("color: #FAFAFA;\n"
"                                            border: none;\n"
"                                            font-size: 11px;\n"
"                                        ")
        self.distorted_images_btn_no.setChecked(True)
        self.distorted_images_btn_no.setObjectName("distorted_images_btn_no")
        self.distorted_images_layout.addWidget(self.distorted_images_btn_no)
        self.verticalLayoutWidget_3 = QtWidgets.QWidget(self.single_file_check)
        self.verticalLayoutWidget_3.setGeometry(QtCore.QRect(690, 300, 160, 80))
        self.verticalLayoutWidget_3.setObjectName("verticalLayoutWidget_3")
        self.all_collisions_layout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget_3)
        self.all_collisions_layout.setContentsMargins(0, 0, 0, 0)
        self.all_collisions_layout.setObjectName("all_collisions_layout")
        self.all_collisions_label = QtWidgets.QLabel(self.verticalLayoutWidget_3)
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(-1)
        font.setBold(True)
        font.setWeight(75)
        self.all_collisions_label.setFont(font)
        self.all_collisions_label.setStyleSheet("border: none;\n"
"                                            color: #FAFAFA;\n"
"                                            font-weight: bold;\n"
"                                            font-size: 12px;\n"
"                                        ")
        self.all_collisions_label.setWordWrap(True)
        self.all_collisions_label.setObjectName("all_collisions_label")
        self.all_collisions_layout.addWidget(self.all_collisions_label)
        self.all_collisions_btn_yes = QtWidgets.QRadioButton(self.verticalLayoutWidget_3)
        self.all_collisions_btn_yes.setStyleSheet("color: #FAFAFA;\n"
"                                            border: none;\n"
"                                            font-size: 11px;\n"
"                                        ")
        self.all_collisions_btn_yes.setObjectName("all_collisions_btn_yes")
        self.all_collisions_layout.addWidget(self.all_collisions_btn_yes)
        self.all_collisions_btn_no = QtWidgets.QRadioButton(self.verticalLayoutWidget_3)
        self.all_collisions_btn_no.setStyleSheet("color: #FAFAFA;\n"
"                                            border: none;\n"
"                                            font-size: 11px;\n"
"                                        ")
        self.all_collisions_btn_no.setChecked(True)
        self.all_collisions_btn_no.setObjectName("all_collisions_btn_no")
        self.all_collisions_layout.addWidget(self.all_collisions_btn_no)
        self.unloading_settings = QtWidgets.QGroupBox(self.single_file_check)
        self.unloading_settings.setGeometry(QtCore.QRect(860, 300, 151, 80))
        self.unloading_settings.setStyleSheet("border: 1px solid #FAFAFA;\n"
"                                color: #FAFAFA;\n"
"                                font-size: 12px;\n"
"                            ")
        self.unloading_settings.setObjectName("unloading_settings")
        self.unloading_btn3 = QtWidgets.QRadioButton(self.unloading_settings)
        self.unloading_btn3.setGeometry(QtCore.QRect(10, 60, 119, 13))
        self.unloading_btn3.setStyleSheet("color: #FAFAFA;\n"
"                                    border: none;\n"
"                                    font-size: 11px;\n"
"                                ")
        self.unloading_btn3.setObjectName("unloading_btn3")
        self.unloading_btn2 = QtWidgets.QRadioButton(self.unloading_settings)
        self.unloading_btn2.setGeometry(QtCore.QRect(10, 39, 119, 13))
        self.unloading_btn2.setStyleSheet("color: #FAFAFA;\n"
"                                    border: none;\n"
"                                    font-size: 11px;\n"
"                                ")
        self.unloading_btn2.setObjectName("unloading_btn2")
        self.unloading_btn1 = QtWidgets.QRadioButton(self.unloading_settings)
        self.unloading_btn1.setGeometry(QtCore.QRect(10, 18, 119, 13))
        self.unloading_btn1.setStyleSheet("color: #FAFAFA;\n"
"                                    border: none;\n"
"                                    font-size: 11px;\n"
"                                ")
        self.unloading_btn1.setObjectName("unloading_btn1")
        self.image_holder = QtWidgets.QLabel(self.single_file_check)
        self.image_holder.setGeometry(QtCore.QRect(690, 40, 320, 160))
        self.image_holder.setStyleSheet("border: 1px dashed #FAFAFA;\n"
"                                border-radius: 2px;\n"
"                            ")
        self.image_holder.setText("")
        self.image_holder.setScaledContents(True)
        self.image_holder.setObjectName("image_holder")
        self.refresh_images = QtWidgets.QPushButton(self.single_file_check)
        self.refresh_images.setGeometry(QtCore.QRect(690, 500, 321, 31))
        self.refresh_images.setStyleSheet("border-radius: 2px;\n"
"                                background-color: #455A64;\n"
"                                border: none;\n"
"                                color: #FAFAFA;\n"
"                                font-size: 12px;\n"
"                                font-weight: bold;\n"
"                            ")
        self.refresh_images.setObjectName("refresh_images")
        self.main_tab_widget.addTab(self.single_file_check, "")
        self.many_files_check = QtWidgets.QWidget()
        self.many_files_check.setStyleSheet("border: 1px solid #FAFAFA;\n"
"                            background-color: #263238;\n"
"                        ")
        self.many_files_check.setObjectName("many_files_check")
        self.main_tab_widget.addTab(self.many_files_check, "")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        self.main_tab_widget.setCurrentIndex(0)
        self.slides_box.setCurrentIndex(1)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Проверка задания ОГЭ"))
        self.input_data.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'Segoe UI\'; font-size:12px; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Загружать файлы можно только по одному. Автоматически проверяются все критерии кроме коллизий. Для проверки коллизий реализован вывод презентаций в виде картинок, строка состояния внизу выводит ответ или ошибки. Конечный файл можно выгрузить в файл формата .xlsx. </span><span style=\" font-size:12pt; font-weight:600;\">Файл можно перетащить в любое место окна</span><span style=\" font-size:12pt;\">.</span><span style=\" font-size:12px;\">                            </span></p></body></html>"))
        self.slides_box.setItemText(self.slides_box.indexOf(self.slide1_page), _translate("MainWindow", "Слайд 2"))
        self.slides_box.setItemText(self.slides_box.indexOf(self.slide2_page), _translate("MainWindow", "Слайд 3"))
        self.statusbar.setPlaceholderText(_translate("MainWindow", "Статусбар"))
        self.get_answer.setText(_translate("MainWindow", "Ввести данные проверки и получить результат"))
        self.image_holder_label.setText(_translate("MainWindow", "Для удобства. Можно загрузить фото"))
        self.txt_img_collisions_label.setText(_translate("MainWindow", "Текст не перекрывает основные изображения"))
        self.txt_img_collisions_btn_yes.setText(_translate("MainWindow", "Да"))
        self.txt_img_collisions_btn_no.setText(_translate("MainWindow", "Нет"))
        self.distorted_images_label.setText(_translate("MainWindow", "Изображения не искажены"))
        self.distorted_images_btn_yes.setText(_translate("MainWindow", "Да"))
        self.distorted_images_btn_no.setText(_translate("MainWindow", "Нет"))
        self.all_collisions_label.setText(_translate("MainWindow", "Не перекрывают текст, заголовок, друг друга"))
        self.all_collisions_btn_yes.setText(_translate("MainWindow", "Да"))
        self.all_collisions_btn_no.setText(_translate("MainWindow", "Нет"))
        self.unloading_settings.setTitle(_translate("MainWindow", "Выгрузка"))
        self.unloading_btn3.setText(_translate("MainWindow", "В статусбар"))
        self.unloading_btn2.setText(_translate("MainWindow", "В .xlsx файл"))
        self.unloading_btn1.setText(_translate("MainWindow", "В .txt файл"))
        self.refresh_images.setText(_translate("MainWindow", "Обновить картинки"))
        self.main_tab_widget.setTabText(self.main_tab_widget.indexOf(self.single_file_check), _translate("MainWindow", "Единичная проверка"))
        self.main_tab_widget.setTabText(self.main_tab_widget.indexOf(self.many_files_check), _translate("MainWindow", "Множественная проверка"))
