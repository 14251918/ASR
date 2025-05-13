import win32com.client
import time

import os.path

from config import settings
from LinkedList import LinkedList
from read_file import read_file


class PresentationSettings:
    def __init__(self, file_path):
        self.ppAPP = None
        self.slide_contexts_json = None
        self.count_slides = None
        self.slideshow = None
        self.presentation_settings = None
        self.presentation = None
        self.file_path = file_path
        if not os.path.isabs(file_path):
            self.file_path = os.path.abspath(file_path)
        self.file_path = self.file_path.replace("\\", "//")
        self.history_slide = LinkedList()

        self.slide_contexts_json = read_file(file_path=file_path)
        self.slide_contexts_str = "\n".join(
            [f"Слайд {slide['slide_number']}: {slide['slide_context']}" for slide in self.slide_contexts_json])

    def run_powerPoint(self):
        if not os.path.exists(self.file_path):
            print(f"Такого файла: {self.file_path} не существует.")
            return
        try:
            self.ppAPP = win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception as e:
            self.ppAPP = win32com.client.Dispatch("PowerPoint.Application")
            self.ppAPP.Visible = True
        try:
            self.presentation = self.ppAPP.Presentations.Open(self.file_path)
            self.presentation_settings = self.presentation.SlideShowSettings
            self.presentation_settings.Run()
            for _ in range(10):
                if self.ppAPP.SlideShowWindows.Count:
                    self.slideshow = self.ppAPP.SlideShowWindows(1)
                    break
                time.sleep(1)
            self.count_slides = self.presentation.Slides.Count

        except Exception as e:
            print(f"Ошибка: {e}. Пожалуйста, укажите полный путь к презентации.")

    def exit_pp(self):
        if self.ppAPP:
            print("PowerPoint закрыт")
            self.ppAPP.DisplayAlerts = False
            self.ppAPP.Quit()
        del self.ppAPP

presentation = PresentationSettings(settings.file_path)
