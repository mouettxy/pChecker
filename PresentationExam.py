# coding: utf-8
import os
import shutil

import win32com.client
from PresentationExamUtils import PresentationExamUtils as Utils
from PresentationExamAnalyze import PresentationExamAnalyze as Analyze
from PresentationExamImages import PresentationExamImages as Images

'''
CHEATSHEET
font_size - Shape.TextFrame.TextRange.Font.Size
font_name - Shape.TextFrame.TextRange.Font.Name
first_line - Shape.TextFrame.TextRange.Lines(1, 1)
id - Shape.Id
'''


class PresentationExam(object):
    def __init__(self, path_to_presentation):
        super().__init__()
        self._path = path_to_presentation
        self._Application = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
        self._Presentation = self._Application.Presentations.Open(self._path, WithWindow=False)
        self._Utils = Utils(self._Application)
        self._Images = Images(self._Presentation, self._Utils)
        self._Analyze = Analyze(self._Presentation, self._Application, self._Utils)

    @property
    def Presentation(self):
        return self._Presentation

    @property
    def Images(self):
        return self._Images

    @property
    def Analyze(self):
        return self._Analyze

    def save(self):
        return self._Presentation.Save()

    def __del__(self):
        if os.path.isdir(os.getcwd() + "\\temp\\"):
            shutil.rmtree(os.getcwd() + "\\temp\\")
        self._Application.Quit()

    def __exit__(self):
        if os.path.isdir(os.getcwd() + "\\temp\\"):
            shutil.rmtree(os.getcwd() + "\\temp\\")
        self._Application.Quit()


PresentationExam = PresentationExam(r"D:\Presentations\002.pptx")
print(PresentationExam.Images.get("skeleton"))
