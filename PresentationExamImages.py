import os
from PIL import ImageDraw, Image


class PresentationExamImages(object):
    def __init__(self, presentation, utils):
        super().__init__()
        self._Presentation = presentation
        self._Utils = utils

    @staticmethod
    def __skeleton_rectangle(draw, shape_dimensions, color, outline):
        return draw.rectangle(
            [shape_dimensions['left'],
             shape_dimensions['top'],
             shape_dimensions['width'] + shape_dimensions['left'],
             shape_dimensions['height'] + shape_dimensions['top']],
            fill=color,
            outline=outline
        )

    def __generate_images_skeleton(self, return_path=False, return_bool=False):
        if not os.path.exists('temp'):
            os.mkdir('temp')
        for Slide in self._Presentation.Slides:
            skeleton = Image.new("RGB",
                                 color="#ffffff",
                                 size=(self._Utils.convert_points_px(self._Presentation.PageSetup.SlideWidth),
                                       self._Utils.convert_points_px(self._Presentation.PageSetup.SlideHeight))
                                 )
            skeleton_path = os.path.abspath(f"temp/skeleton_{Slide.SlideIndex}.jpg")
            skeleton_draw = ImageDraw.Draw(skeleton)
            for Shape in Slide.Shapes:
                shape_dimensions = self._Utils.get_shape_dimensions(Shape)
                if self._Utils.is_text(Shape) is True:
                    self.__skeleton_rectangle(skeleton_draw, shape_dimensions, "yellow", "red")
                elif self._Utils.is_text(Shape) is None:
                    self.__skeleton_rectangle(skeleton_draw, shape_dimensions, "orange", "yellow")
                elif self._Utils.is_image(Shape) is True:
                    self.__skeleton_rectangle(skeleton_draw, shape_dimensions, "blue", "yellow")
                else:
                    self.__skeleton_rectangle(skeleton_draw, shape_dimensions, "red", "yellow")
                skeleton.save(skeleton_path)
            if return_path is True:
                yield skeleton_path
        if return_bool is True:
            return True

    def __generate_images_screenshots(self, return_path=False, return_bool=False):
        if not os.path.exists('temp'):
            os.mkdir('temp')
        for Slide in self._Presentation.Slides:
            screenshot_path = os.path.abspath(f"temp/slide_{Slide.SlideIndex}.jpg")
            Slide.Export(screenshot_path, "JPG")
            if return_path is True:
                yield screenshot_path
        if return_bool:
            return True

    def get(self, typeof="exact", return_path=True):
        if typeof == "exact":
            if return_path:
                return list(self.__generate_images_screenshots(return_path=True))
            else:
                return self.__generate_images_screenshots(return_bool=True)
        elif typeof == "skeleton":
            if return_path:
                return list(self.__generate_images_skeleton(return_path=True))
            else:
                return self.__generate_images_skeleton(return_bool=True)
        else:
            return "Mismatched type of generation"
