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
        MainWindow.resize(1020, 672)
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
        self.single_file_check.setStyleSheet("background-color: #212121;")
        self.single_file_check.setObjectName("single_file_check")
        self.input_data = QtWidgets.QTextEdit(self.single_file_check)
        self.input_data.setGeometry(QtCore.QRect(10, 10, 671, 81))
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(-1)
        self.input_data.setFont(font)
        self.input_data.setStyleSheet("color: #FAFAFA;\n"
"font-size: 10px;\n"
"background-color: #424242;\n"
"border: none;\n"
"border-radius: 2px;\n"
"                            ")
        self.input_data.setObjectName("input_data")
        self.slides_box = QtWidgets.QToolBox(self.single_file_check)
        self.slides_box.setGeometry(QtCore.QRect(10, 100, 671, 451))
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
        self.statusbar.setGeometry(QtCore.QRect(10, 560, 671, 81))
        self.statusbar.setStyleSheet("border: none;\n"
"padding-left: 4px;\n"
"padding-right: 4px;\n"
"color: #fafafa;\n"
"background-color: #424242;\n"
"border-radius: 2px;\n"
"font-weight: bold;\n"
"                            ")
        self.statusbar.setObjectName("statusbar")
        self.get_answer = QtWidgets.QPushButton(self.single_file_check)
        self.get_answer.setGeometry(QtCore.QRect(690, 610, 321, 31))
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(-1)
        font.setBold(True)
        font.setWeight(75)
        self.get_answer.setFont(font)
        self.get_answer.setCursor(QtGui.QCursor(QtCore.Qt.ClosedHandCursor))
        self.get_answer.setStyleSheet("QPushButton {\n"
"    border-radius: 2px;\n"
"    background-color: #424242;\n"
"    border: none;\n"
"    color: #FAFAFA;\n"
"    font-size: 12px;\n"
"    font-weight: bold;\n"
"}\n"
"QPushButton:hover {\n"
"    background-color: #757575;\n"
"}\n"
"QPushButton:clicked {\n"
"    background-color: #9E9E9E;\n"
"}\n"
"\n"
"                            ")
        self.get_answer.setObjectName("get_answer")
        self.groupBox = QtWidgets.QGroupBox(self.single_file_check)
        self.groupBox.setGeometry(QtCore.QRect(690, 240, 321, 311))
        self.groupBox.setStyleSheet("QGroupBox {\n"
"border-color: #FAFAFA;\n"
"background-color:#616161;\n"
"font-size: 14px;\n"
"border-radius: 2px;\n"
"}\n"
"QGroupBox::title {\n"
"font-weight: bold;\n"
"border-top-left-radius: 2px;\n"
"border-top-right-radius: 2px;\n"
"padding: 2px 75px;\n"
"background-color: #424242;\n"
"color: #FAFAFA;\n"
"}")
        self.groupBox.setObjectName("groupBox")
        self.content_compliace_frame = QtWidgets.QFrame(self.groupBox)
        self.content_compliace_frame.setGeometry(QtCore.QRect(10, 235, 301, 65))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.content_compliace_frame.sizePolicy().hasHeightForWidth())
        self.content_compliace_frame.setSizePolicy(sizePolicy)
        self.content_compliace_frame.setStyleSheet("background: #757575;\n"
"border-radius: 2px;")
        self.content_compliace_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.content_compliace_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.content_compliace_frame.setObjectName("content_compliace_frame")
        self.gridLayout_4 = QtWidgets.QGridLayout(self.content_compliace_frame)
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.content_compliance_btn_yes = QtWidgets.QRadioButton(self.content_compliace_frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.content_compliance_btn_yes.sizePolicy().hasHeightForWidth())
        self.content_compliance_btn_yes.setSizePolicy(sizePolicy)
        self.content_compliance_btn_yes.setStyleSheet("color: #FAFAFA;\n"
"border: none;\n"
"background: #616161;\n"
"font-size: 11px;\n"
"padding: 5px;")
        self.content_compliance_btn_yes.setObjectName("content_compliance_btn_yes")
        self.gridLayout_4.addWidget(self.content_compliance_btn_yes, 0, 1, 1, 1)
        self.content_compliance_btn_no = QtWidgets.QRadioButton(self.content_compliace_frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.content_compliance_btn_no.sizePolicy().hasHeightForWidth())
        self.content_compliance_btn_no.setSizePolicy(sizePolicy)
        self.content_compliance_btn_no.setStyleSheet("color: #FAFAFA;\n"
"border: none;\n"
"background: #616161;\n"
"font-size: 11px;\n"
"padding: 5px;")
        self.content_compliance_btn_no.setChecked(True)
        self.content_compliance_btn_no.setObjectName("content_compliance_btn_no")
        self.gridLayout_4.addWidget(self.content_compliance_btn_no, 1, 1, 1, 1)
        self.content_compliance_label = QtWidgets.QLabel(self.content_compliace_frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.content_compliance_label.sizePolicy().hasHeightForWidth())
        self.content_compliance_label.setSizePolicy(sizePolicy)
        self.content_compliance_label.setMinimumSize(QtCore.QSize(200, 0))
        self.content_compliance_label.setMaximumSize(QtCore.QSize(200, 16777215))
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(-1)
        font.setBold(False)
        font.setWeight(50)
        self.content_compliance_label.setFont(font)
        self.content_compliance_label.setStyleSheet("border: none;\n"
"color: #FAFAFA;\n"
"background: #616161;\n"
"font-weight: normal;\n"
"font-size: 12px;\n"
"padding: 5px;")
        self.content_compliance_label.setWordWrap(True)
        self.content_compliance_label.setObjectName("content_compliance_label")
        self.gridLayout_4.addWidget(self.content_compliance_label, 0, 0, 2, 1)
        self.txt_images_collisions_frame = QtWidgets.QFrame(self.groupBox)
        self.txt_images_collisions_frame.setGeometry(QtCore.QRect(10, 165, 301, 65))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.txt_images_collisions_frame.sizePolicy().hasHeightForWidth())
        self.txt_images_collisions_frame.setSizePolicy(sizePolicy)
        self.txt_images_collisions_frame.setStyleSheet("background: #757575;\n"
"border-radius: 2px;")
        self.txt_images_collisions_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.txt_images_collisions_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.txt_images_collisions_frame.setObjectName("txt_images_collisions_frame")
        self.gridLayout_3 = QtWidgets.QGridLayout(self.txt_images_collisions_frame)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.txt_img_collisions_btn_no = QtWidgets.QRadioButton(self.txt_images_collisions_frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.txt_img_collisions_btn_no.sizePolicy().hasHeightForWidth())
        self.txt_img_collisions_btn_no.setSizePolicy(sizePolicy)
        self.txt_img_collisions_btn_no.setStyleSheet("color: #FAFAFA;\n"
"border: none;\n"
"background: #616161;\n"
"font-size: 11px;\n"
"padding: 5px;")
        self.txt_img_collisions_btn_no.setChecked(True)
        self.txt_img_collisions_btn_no.setObjectName("txt_img_collisions_btn_no")
        self.gridLayout_3.addWidget(self.txt_img_collisions_btn_no, 1, 2, 1, 1)
        self.txt_img_collisions_btn_yes = QtWidgets.QRadioButton(self.txt_images_collisions_frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.txt_img_collisions_btn_yes.sizePolicy().hasHeightForWidth())
        self.txt_img_collisions_btn_yes.setSizePolicy(sizePolicy)
        self.txt_img_collisions_btn_yes.setStyleSheet("color: #FAFAFA;\n"
"border: none;\n"
"background: #616161;\n"
"font-size: 11px;\n"
"padding: 5px;")
        self.txt_img_collisions_btn_yes.setObjectName("txt_img_collisions_btn_yes")
        self.gridLayout_3.addWidget(self.txt_img_collisions_btn_yes, 0, 2, 1, 1)
        self.txt_img_collisions_label = QtWidgets.QLabel(self.txt_images_collisions_frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.txt_img_collisions_label.sizePolicy().hasHeightForWidth())
        self.txt_img_collisions_label.setSizePolicy(sizePolicy)
        self.txt_img_collisions_label.setMinimumSize(QtCore.QSize(200, 0))
        self.txt_img_collisions_label.setMaximumSize(QtCore.QSize(200, 16777215))
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(-1)
        font.setBold(False)
        font.setWeight(50)
        self.txt_img_collisions_label.setFont(font)
        self.txt_img_collisions_label.setStyleSheet("border: none;\n"
"color: #FAFAFA;\n"
"background: #616161;\n"
"font-weight: normal;\n"
"font-size: 12px;\n"
"padding: 5px;")
        self.txt_img_collisions_label.setWordWrap(True)
        self.txt_img_collisions_label.setObjectName("txt_img_collisions_label")
        self.gridLayout_3.addWidget(self.txt_img_collisions_label, 0, 1, 2, 1)
        self.distorted_images_frame = QtWidgets.QFrame(self.groupBox)
        self.distorted_images_frame.setGeometry(QtCore.QRect(10, 95, 301, 65))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.distorted_images_frame.sizePolicy().hasHeightForWidth())
        self.distorted_images_frame.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(9)
        font.setBold(False)
        font.setWeight(50)
        self.distorted_images_frame.setFont(font)
        self.distorted_images_frame.setStyleSheet("background: #757575;\n"
"border-radius: 2px;")
        self.distorted_images_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.distorted_images_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.distorted_images_frame.setObjectName("distorted_images_frame")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.distorted_images_frame)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.distorted_images_btn_no = QtWidgets.QRadioButton(self.distorted_images_frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.distorted_images_btn_no.sizePolicy().hasHeightForWidth())
        self.distorted_images_btn_no.setSizePolicy(sizePolicy)
        self.distorted_images_btn_no.setMinimumSize(QtCore.QSize(80, 0))
        self.distorted_images_btn_no.setMaximumSize(QtCore.QSize(80, 16777215))
        self.distorted_images_btn_no.setStyleSheet("color: #FAFAFA;\n"
"border: none;\n"
"background: #616161;\n"
"font-size: 11px;\n"
"padding: 5px;")
        self.distorted_images_btn_no.setChecked(True)
        self.distorted_images_btn_no.setObjectName("distorted_images_btn_no")
        self.gridLayout_2.addWidget(self.distorted_images_btn_no, 1, 1, 1, 1)
        self.distorted_images_btn_yes = QtWidgets.QRadioButton(self.distorted_images_frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.distorted_images_btn_yes.sizePolicy().hasHeightForWidth())
        self.distorted_images_btn_yes.setSizePolicy(sizePolicy)
        self.distorted_images_btn_yes.setMinimumSize(QtCore.QSize(80, 0))
        self.distorted_images_btn_yes.setMaximumSize(QtCore.QSize(80, 16777215))
        self.distorted_images_btn_yes.setStyleSheet("color: #FAFAFA;\n"
"border: none;\n"
"background: #616161;\n"
"font-size: 11px;\n"
"padding: 5px;")
        self.distorted_images_btn_yes.setObjectName("distorted_images_btn_yes")
        self.gridLayout_2.addWidget(self.distorted_images_btn_yes, 0, 1, 1, 1)
        self.distorted_images_label = QtWidgets.QLabel(self.distorted_images_frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.distorted_images_label.sizePolicy().hasHeightForWidth())
        self.distorted_images_label.setSizePolicy(sizePolicy)
        self.distorted_images_label.setMinimumSize(QtCore.QSize(200, 0))
        self.distorted_images_label.setMaximumSize(QtCore.QSize(200, 16777215))
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(-1)
        font.setBold(False)
        font.setWeight(50)
        self.distorted_images_label.setFont(font)
        self.distorted_images_label.setStyleSheet("border: none;\n"
"color: #FAFAFA;\n"
"background: #616161;\n"
"font-weight: normal;\n"
"font-size: 12px;\n"
"padding: 5px;")
        self.distorted_images_label.setWordWrap(True)
        self.distorted_images_label.setObjectName("distorted_images_label")
        self.gridLayout_2.addWidget(self.distorted_images_label, 0, 0, 2, 1)
        self.all_collisions_frame = QtWidgets.QFrame(self.groupBox)
        self.all_collisions_frame.setGeometry(QtCore.QRect(10, 25, 301, 65))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.all_collisions_frame.sizePolicy().hasHeightForWidth())
        self.all_collisions_frame.setSizePolicy(sizePolicy)
        self.all_collisions_frame.setStyleSheet("background: #757575;\n"
"border-radius: 2px;")
        self.all_collisions_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.all_collisions_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.all_collisions_frame.setObjectName("all_collisions_frame")
        self.gridLayout = QtWidgets.QGridLayout(self.all_collisions_frame)
        self.gridLayout.setObjectName("gridLayout")
        self.all_collisions_btn_no = QtWidgets.QRadioButton(self.all_collisions_frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.all_collisions_btn_no.sizePolicy().hasHeightForWidth())
        self.all_collisions_btn_no.setSizePolicy(sizePolicy)
        self.all_collisions_btn_no.setMinimumSize(QtCore.QSize(80, 0))
        self.all_collisions_btn_no.setMaximumSize(QtCore.QSize(80, 16777215))
        self.all_collisions_btn_no.setStyleSheet("color: #FAFAFA;\n"
"border: none;\n"
"background: #616161;\n"
"font-size: 11px;\n"
"padding: 5px;")
        self.all_collisions_btn_no.setChecked(True)
        self.all_collisions_btn_no.setObjectName("all_collisions_btn_no")
        self.gridLayout.addWidget(self.all_collisions_btn_no, 1, 1, 1, 1)
        self.all_collisions_btn_yes = QtWidgets.QRadioButton(self.all_collisions_frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.all_collisions_btn_yes.sizePolicy().hasHeightForWidth())
        self.all_collisions_btn_yes.setSizePolicy(sizePolicy)
        self.all_collisions_btn_yes.setMinimumSize(QtCore.QSize(80, 0))
        self.all_collisions_btn_yes.setMaximumSize(QtCore.QSize(80, 16777215))
        self.all_collisions_btn_yes.setStyleSheet("color: #FAFAFA;\n"
"border: none;\n"
"background: #616161;\n"
"font-size: 11px;\n"
"padding: 5px;")
        self.all_collisions_btn_yes.setObjectName("all_collisions_btn_yes")
        self.gridLayout.addWidget(self.all_collisions_btn_yes, 0, 1, 1, 1)
        self.all_collisions_label = QtWidgets.QLabel(self.all_collisions_frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.all_collisions_label.sizePolicy().hasHeightForWidth())
        self.all_collisions_label.setSizePolicy(sizePolicy)
        self.all_collisions_label.setMinimumSize(QtCore.QSize(200, 0))
        self.all_collisions_label.setMaximumSize(QtCore.QSize(200, 16777215))
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(-1)
        font.setBold(False)
        font.setWeight(50)
        self.all_collisions_label.setFont(font)
        self.all_collisions_label.setStyleSheet("border: none;\n"
"color: #FAFAFA;\n"
"background: #616161;\n"
"font-weight: normal;\n"
"font-size: 12px;\n"
"padding: 5px;")
        self.all_collisions_label.setWordWrap(True)
        self.all_collisions_label.setObjectName("all_collisions_label")
        self.gridLayout.addWidget(self.all_collisions_label, 0, 0, 2, 1)
        self.image_holder_box_title = QtWidgets.QGroupBox(self.single_file_check)
        self.image_holder_box_title.setGeometry(QtCore.QRect(690, 10, 321, 221))
        self.image_holder_box_title.setStyleSheet("QGroupBox {\n"
"border-color: #FAFAFA;\n"
"background-color:#616161;\n"
"font-size: 14px;\n"
"border-radius: 2px;\n"
"}\n"
"QGroupBox::title {\n"
"font-weight: bold;\n"
"border-top-left-radius: 2px;\n"
"border-top-right-radius: 2px;\n"
"padding: 2px 34px;\n"
"background-color: #424242;\n"
"color: #FAFAFA;\n"
"}")
        self.image_holder_box_title.setObjectName("image_holder_box_title")
        self.image_holder = QtWidgets.QLabel(self.image_holder_box_title)
        self.image_holder.setGeometry(QtCore.QRect(10, 30, 301, 181))
        self.image_holder.setStyleSheet("border: 1px dashed #FAFAFA;\n"
"background: #424242;\n"
"border-radius: 2px;\n"
"                            ")
        self.image_holder.setText("")
        self.image_holder.setScaledContents(True)
        self.image_holder.setObjectName("image_holder")
        self.unloading_settings = QtWidgets.QGroupBox(self.single_file_check)
        self.unloading_settings.setGeometry(QtCore.QRect(690, 560, 321, 81))
        self.unloading_settings.setStyleSheet("QGroupBox {\n"
"border-color: #FAFAFA;\n"
"background-color:#616161;\n"
"font-size: 14px;\n"
"border-radius: 2px;\n"
"}\n"
"QGroupBox::title {\n"
"border-top-left-radius: 2px;\n"
"border-top-right-radius: 2px;\n"
"padding: 2px 130px;\n"
"background-color: #424242;\n"
"color: #FAFAFA;\n"
"}\n"
"                            ")
        self.unloading_settings.setObjectName("unloading_settings")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.unloading_settings)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.unloading_btn3 = QtWidgets.QRadioButton(self.unloading_settings)
        self.unloading_btn3.setStyleSheet("color: #FAFAFA;\n"
"background-color: #616161;\n"
"border: none;\n"
"font-size: 12px;\n"
"margin-bottom: 10px;\n"
"                                ")
        self.unloading_btn3.setObjectName("unloading_btn3")
        self.horizontalLayout_2.addWidget(self.unloading_btn3)
        self.unloading_btn2 = QtWidgets.QRadioButton(self.unloading_settings)
        self.unloading_btn2.setStyleSheet("color: #FAFAFA;\n"
"background-color: #616161;\n"
"border: none;\n"
"font-size: 12px;\n"
"margin-bottom: 10px;\n"
"                                ")
        self.unloading_btn2.setObjectName("unloading_btn2")
        self.horizontalLayout_2.addWidget(self.unloading_btn2)
        self.unloading_btn1 = QtWidgets.QRadioButton(self.unloading_settings)
        self.unloading_btn1.setStyleSheet("color: #FAFAFA;\n"
"background-color: #616161;\n"
"border: none;\n"
"font-size: 12px;\n"
"margin-bottom: 10px;\n"
"                                ")
        self.unloading_btn1.setObjectName("unloading_btn1")
        self.horizontalLayout_2.addWidget(self.unloading_btn1)
        self.input_data.raise_()
        self.slides_box.raise_()
        self.statusbar.raise_()
        self.groupBox.raise_()
        self.image_holder_box_title.raise_()
        self.unloading_settings.raise_()
        self.get_answer.raise_()
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
"</style></head><body style=\" font-family:\'Segoe UI\'; font-size:10px; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:10pt;\">Загружать файлы можно только по одному. Автоматически проверяются все критерии кроме коллизий. Для проверки коллизий реализован вывод презентаций в виде картинок, строка состояния внизу выводит ответ или ошибки. Конечный файл можно выгрузить в файл формата .xlsx. </span><span style=\" font-size:10pt; font-weight:600;\">Файл можно перетащить в любое место окна</span><span style=\" font-size:10pt;\">.                            </span></p></body></html>"))
        self.slides_box.setItemText(self.slides_box.indexOf(self.slide1_page), _translate("MainWindow", "Слайд 2"))
        self.slides_box.setItemText(self.slides_box.indexOf(self.slide2_page), _translate("MainWindow", "Слайд 3"))
        self.statusbar.setPlaceholderText(_translate("MainWindow", "Статусбар"))
        self.get_answer.setText(_translate("MainWindow", "Ввести данные проверки и получить результат"))
        self.groupBox.setTitle(_translate("MainWindow", "Элементы ручной проверки"))
        self.content_compliance_btn_yes.setText(_translate("MainWindow", "Да"))
        self.content_compliance_btn_no.setText(_translate("MainWindow", "Нет"))
        self.content_compliance_label.setText(_translate("MainWindow", "Соответствует теме презентации"))
        self.txt_img_collisions_btn_no.setText(_translate("MainWindow", "Нет"))
        self.txt_img_collisions_btn_yes.setText(_translate("MainWindow", "Да"))
        self.txt_img_collisions_label.setText(_translate("MainWindow", "Текст не перекрывает основные изображения"))
        self.distorted_images_btn_no.setText(_translate("MainWindow", "Нет"))
        self.distorted_images_btn_yes.setText(_translate("MainWindow", "Да"))
        self.distorted_images_label.setText(_translate("MainWindow", "Изображения не искажены"))
        self.all_collisions_btn_no.setText(_translate("MainWindow", "Нет"))
        self.all_collisions_btn_yes.setText(_translate("MainWindow", "Да"))
        self.all_collisions_label.setText(_translate("MainWindow", "Изображения не перекрывают текст, себя, элементы"))
        self.image_holder_box_title.setTitle(_translate("MainWindow", "Для удобства: Можно загрузить фото"))
        self.unloading_settings.setTitle(_translate("MainWindow", "Выгрузка"))
        self.unloading_btn3.setText(_translate("MainWindow", "В статусбар"))
        self.unloading_btn2.setText(_translate("MainWindow", "В .xlsx файл"))
        self.unloading_btn1.setText(_translate("MainWindow", "В .txt файл"))
        self.main_tab_widget.setTabText(self.main_tab_widget.indexOf(self.single_file_check), _translate("MainWindow", "Единичная проверка"))
        self.main_tab_widget.setTabText(self.main_tab_widget.indexOf(self.many_files_check), _translate("MainWindow", "Множественная проверка"))
