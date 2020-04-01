# coding: utf-8
import shutil
from pathlib import Path

import win32com.client

from PresentationExamAnalyze import PresentationExamAnalyze as Analyze
from PresentationExamImages import PresentationExamImages as Images
from PresentationExamUtils import PresentationExamUtils as Utils


def open_presentation(path):
    win32com.client.Dispatch("PowerPoint.Application").Presentations.Open(path)


def upload_images(from_directory):
    f_dir = Path(from_directory).resolve()
    if f_dir.is_dir():
        path = Path('original_images').resolve()
        if path.exists():
            shutil.rmtree(path, ignore_errors=True)
        path.mkdir(parents=True)
        for filename in f_dir.iterdir():
            if filename.suffix in ['.png', '.jpg', '.jpeg']:
                shutil.copy(filename, path, follow_symlinks=True)
        return True
    return False  # TODO generate expression here


class PresentationExam(object):
    def __init__(self, path_to_presentation):
        super().__init__()
        self._path = path_to_presentation
        self._Application = win32com.client.Dispatch("PowerPoint.Application")
        self._Presentation = self._Application.Presentations.Open(self._path, WithWindow=False)
        self._Utils = Utils(self._Application)
        self._Images = Images(self._Presentation, self._Utils)
        self._Analyze = Analyze(self._Presentation, self._Application, self._Utils, self._Images)

    @property
    def Utils(self):
        return self._Utils

    @property
    def Presentation(self):
        return self._Presentation

    @property
    def Images(self):
        return self._Images

    @property
    def Analyze(self):
        return self._Analyze

    def __del__(self):
        self._Application.Quit()

    def __exit__(self):
        self._Application.Quit()
