import os
from CheckPresentation import Analyze, Warnings, Images
import win32com.client
from pp_classes import MSOPPT, MSO
import time


class Testing(Analyze.Analyze, Warnings.Warnings, Images.Images):
    """
    Класс для реализации каких либо новых функций. Создан для удобства отладки и разработки.
    """

    @staticmethod
    def prs_list(self, directory="D:\\Presentations"):
        return [directory + "\\" + d for d in os.listdir(dir)]

    def generate_slide_images(self):
        app = win32com.client.Dispatch("PowerPoint.Application")
        prs = app.Presentations.Open(self.path_to_presentation, WithWindow=False)
        directory = os.path.abspath(os.getcwd())
        counter = 1
        result = []
        for s in prs.Slides:
            path = f"{directory}\\slide_images\\slide_{counter}.jpg"
            s.Export(path, "JPG")
            result.append(path)
            counter += 1
        '''
        Этот фрагмент показывает как я пытался починить рамки shape'оф программно. Безуспешно пока что.
        for sld in prs.Slides:
            for shp in sld.Shapes:
                if shp.HasTextFrame:
                    if not shp.TextFrame.AutoSize == MSOPPT.constants.ppAutoSizeShapeToFitText:
                        shp.TextFrame.WordWrap = MSO.constants.msoFalse
                        shp.TextFrame.AutoSize = MSOPPT.constants.ppAutoSizeShapeToFitText
                        shp.TextFrame.WordWrap = MSO.constants.msoTrue
        prs.Save()
        '''
        return result