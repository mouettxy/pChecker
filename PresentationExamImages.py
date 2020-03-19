import inspect
import os
import shutil
import zipfile

import cv2
import imagehash
import numpy as np
from PIL import ImageDraw, Image

from MSOCONSTANTS import ppShapeFormatJPG, msoTrue
from PresentationExamLayouts import PresentationExamLayouts as Layouts


class PresentationExamImages(object):
    def __init__(self, presentation, utils):
        super().__init__()
        self._Presentation = presentation
        self._Utils = utils
        self._layouts = inspect.getmembers(Layouts, predicate=inspect.isfunction)

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

    def __generate_images_skeleton(self, return_path=False, return_bool=False, directory="temp"):
        if not os.path.exists(directory):
            os.mkdir(directory)
        for Slide in self._Presentation.Slides:
            skeleton = Image.new("RGB",
                                 color="#ffffff",
                                 size=(self._Utils.convert_points_px(self._Presentation.PageSetup.SlideWidth),
                                       self._Utils.convert_points_px(self._Presentation.PageSetup.SlideHeight))
                                 )
            skeleton_path = os.path.abspath(f"{directory}/skeleton_{Slide.SlideIndex}.jpg")
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

    def __generate_images_layout(self, lt="layout_1", directory="temp"):
        """
        Experimental
        :return: None
        """
        if not os.path.exists(directory):
            os.mkdir(directory)
        width = self._Utils.convert_points_px(self._Presentation.PageSetup.SlideWidth)
        height = self._Utils.convert_points_px(self._Presentation.PageSetup.SlideHeight)
        result = []
        layout = self._layouts[lt][1](width, height)
        for slide in layout:
            image_path = os.path.abspath(f"{directory}/image_{slide}.png")
            cv2.imwrite(image_path, 255 * np.ones((height, width, 3), np.uint8))
            image = cv2.imread(image_path)
            overlay = image.copy()
            for i in layout[slide]['images']:
                l, t, w, h = int(i[0]), int(i[1]), int(i[2]), int(i[3])
                cv2.rectangle(overlay, (l, t), (w + l, h + t), (0, 255, 0), -1)
            for t in layout[slide]['text_blocks']:
                l, t, w, h = int(t[0]), int(t[1]), int(t[2]), int(t[3])
                cv2.rectangle(overlay, (l, t), (w + l, h + t), (255, 0, 0), -1)
            cv2.addWeighted(overlay, 0.7, image, 0.3, 0, image)
            cv2.imwrite(image_path, image)
            result.append(image_path)
        return result

    def __get_picture_shape_images(self, return_path=False, return_bool=False):
        """
        Reserved for internal use
        """
        if not os.path.exists('shapes_images'):
            os.mkdir('shapes_images')
        counter = 0
        for Slide in self._Presentation.Slides:
            for Shape in Slide.Shapes:
                if self._Utils.is_image(Shape):
                    counter += 1
                    path = os.getcwd() + f"\\shapes_images\\picture_{counter}.jpg"
                    Shape.Export(path, ppShapeFormatJPG)
                    if return_path:
                        yield path
        if return_bool:
            return True

    def __save_original_images_presentation(self, directory="original_images_presentation"):
        if not os.path.exists(directory):
            os.mkdir(directory)
        file, c = zipfile.ZipFile(os.path.join(self._Presentation.Path, self._Presentation.Name)), 0
        for f in file.namelist():
            if f.startswith('ppt/media'):
                save_path = os.path.abspath(f"{directory}/image_{c}.jpg")
                shutil.copyfileobj(file.open(f), open(save_path, "wb"))
                c += 1
                yield save_path

    def __compare_images(self, directory="original_images"):
        path_to_shape_images = list(self.__save_original_images_presentation())
        compare_images_counter = 0
        for original_image_name in os.listdir(os.path.abspath(directory)):
            for shape_image_path in path_to_shape_images:
                original_image = Image.open(os.path.abspath(f"{directory}/{original_image_name}"))
                shape_image = Image.open(shape_image_path)
                print(imagehash.average_hash(original_image), imagehash.average_hash(shape_image))
                if imagehash.average_hash(original_image) == imagehash.average_hash(shape_image):
                    compare_images_counter += 1
                    break
        else:
            if len(path_to_shape_images) <= 2:
                return True
            return False

    def __generate_images_screenshots(self, return_path=False, return_bool=False, directory="temp"):
        if not os.path.exists(directory):
            os.mkdir(directory)
        for Slide in self._Presentation.Slides:
            screenshot_path = os.path.abspath(f"{directory}/slide_{Slide.SlideIndex}.jpg")
            Slide.Export(screenshot_path, "JPG")
            if return_path is True:
                yield screenshot_path
        if return_bool:
            return True

    @staticmethod
    def __get_shape_percentage_width_height(Shape):
        shape_width, shape_height = Shape.Width, Shape.Height
        Shape.ScaleWidth(1, msoTrue)
        Shape.ScaleHeight(1, msoTrue)
        return round(shape_width / Shape.Width * 100), round(shape_height / Shape.Height * 100)

    def distorted_images(self):
        for Slide in self._Presentation.Slides:
            for Shape in Slide.Shapes:
                if not self._Utils.is_text(Shape):
                    w, h = self.__get_shape_percentage_width_height(Shape)
                    if abs(w - h) > 10:
                        return True
        return False

    def compare(self, directory="original_images"):
        if os.path.exists(os.path.abspath(directory)):
            return self.__compare_images()
        else:
            return "Не загружены изображения."

    def get(self, typeof="exact_match", return_path=True, lt="layout_1"):
        if typeof == "exact_match":
            if return_path:
                return list(self.__generate_images_screenshots(return_path=True))
            else:
                return self.__generate_images_screenshots(return_bool=True)
        elif typeof == "skeleton":
            if return_path:
                return list(self.__generate_images_skeleton(return_path=True))
            else:
                return self.__generate_images_skeleton(return_bool=True)
        elif typeof == "layout":
            return self.__generate_images_layout(lt)
        else:
            return "Mismatched type of generation"
