import inspect
import shutil
import zipfile
from pathlib import Path

import imagehash
from PIL import ImageDraw, Image

from MSOCONSTANTS import ppShapeFormatJPG
from PresentationExamLayouts import PresentationExamLayouts as Layouts


class PresentationExamImages(object):
    def __init__(self, presentation, utils):
        super().__init__()
        self._Presentation = presentation
        self._Utils = utils
        self._layouts = inspect.getmembers(Layouts, predicate=inspect.isfunction)
        self.path = Path(self._Presentation.Path + "/" + self._Presentation.Name).resolve()
        self.destination = Path(f"temp/{self._Presentation.Name}").resolve()
        self.destination.mkdir(exist_ok=True, parents=True)

    @staticmethod
    def __draw_rectangle(draw, shape_dimensions, color, outline="red"):
        return draw.rectangle(
            [shape_dimensions['left'],
             shape_dimensions['top'],
             shape_dimensions['width'] + shape_dimensions['left'],
             shape_dimensions['height'] + shape_dimensions['top']],
            fill=color,
            outline=outline
        )

    def __generate_images_skeleton(self):
        paths = []
        for Slide in self._Presentation.Slides:
            path = Path.joinpath(self.destination, f"skeleton_{Slide.SlideIndex}.jpg")
            image = Image.new(
                "RGB",
                color="white",
                size=(self._Utils.convert_points_px(self._Presentation.PageSetup.SlideWidth),
                      self._Utils.convert_points_px(self._Presentation.PageSetup.SlideHeight))
            )
            skeleton_draw = ImageDraw.Draw(image)
            for Shape in Slide.Shapes:
                shape_dimensions = self._Utils.get_shape_dimensions(Shape)
                if self._Utils.is_text(Shape) is True:
                    self.__draw_rectangle(skeleton_draw, shape_dimensions, "yellow", "red")
                elif self._Utils.is_text(Shape) is None:
                    self.__draw_rectangle(skeleton_draw, shape_dimensions, "orange", "yellow")
                elif self._Utils.is_image(Shape) is True:
                    self.__draw_rectangle(skeleton_draw, shape_dimensions, "blue", "yellow")
                else:
                    self.__draw_rectangle(skeleton_draw, shape_dimensions, "red", "yellow")
            image.save(path)
            paths.append(path)
        return paths

    def __generate_images_layout(self, lt="layout_1"):
        """
        Experimental
        """
        paths = []
        color = (250, 250, 250, 1)
        layout = None
        for layout in self._layouts:
            if layout[0] == lt:
                layout = layout[1](self._Utils.convert_points_px(self._Presentation.PageSetup.SlideWidth),
                                   self._Utils.convert_points_px(self._Presentation.PageSetup.SlideHeight))
                break
        for slide in layout:
            path = Path.joinpath(self.destination, f"layout_{slide}.png")
            image = Image.new(
                "RGB",
                (self._Utils.convert_points_px(self._Presentation.PageSetup.SlideWidth),
                 self._Utils.convert_points_px(self._Presentation.PageSetup.SlideHeight)),
                "white"
            )
            draw = ImageDraw.Draw(image, "RGBA")
            for block_type in layout[slide]:
                if block_type == "title":
                    color = (27, 94, 32, 100)
                elif block_type == "images":
                    color = (245, 127, 23, 120)
                elif block_type == "text_blocks":
                    color = (26, 35, 126, 175)
                for dims in layout[slide][block_type]:
                    dims = {'left': dims[0], 'top': dims[1], 'width': dims[2], 'height': dims[3]}
                    self.__draw_rectangle(draw, dims, color)
            image.save(path)
            paths.append(path)
        return paths

    def __get_picture_shape_images(self):
        """
        Reserved for internal use
        """
        destination = Path.joinpath(self.destination, 'shapes')
        destination.mkdir(exist_ok=True, parents=True)
        paths = []
        for Slide in self._Presentation.Slides:
            for Shape in Slide.Shapes:
                if self._Utils.is_image(Shape):
                    path = Path.joinpath(destination, f"{Slide.SlideIndex}_{Shape.Id}.jpg")
                    Shape.Export(path, ppShapeFormatJPG)
                    paths.append(path)
        return paths

    def __save_original_images_presentation(self):
        paths = []
        file = zipfile.ZipFile(self.path)
        destination = Path.joinpath(self.destination, "media")
        destination.mkdir(parents=True, exist_ok=True)
        for f in file.namelist():
            if f.startswith('ppt/media'):
                path = Path.joinpath(destination, Path(f).name)
                shutil.copyfileobj(file.open(f), open(path, "wb"))
                paths.append(path)
        return paths

    def __compare_images(self, path='original_images'):
        original_images, counter, shape_images = [], 0, self.__save_original_images_presentation()
        for extension in ['*.png', '*.jpg', '*.jpeg']:
            original_images.extend(Path(path).resolve().glob(extension))
        for o_path in original_images:
            for s_path in shape_images:
                o_image = Image.open(o_path)
                s_image = Image.open(s_path)
                if imagehash.average_hash(o_image) == imagehash.average_hash(s_image):
                    counter += 1
                    break
        if len(shape_images) == counter:
            return True
        return False

    def __generate_images_screenshots(self, thumb=False):
        paths = []
        for Slide in self._Presentation.Slides:
            path = Path.joinpath(self.destination, f"screenshot_{Slide.SlideIndex}.jpg")
            Slide.Export(path, "JPG")
            paths.append(path)
            if thumb:
                image = Image.open(path)
                image.thumbnail((200, 200), Image.ANTIALIAS)
                image.save(Path(self.destination, "thumb.jpg"))
                return Path(self.destination, "thumb.jpg")
        return paths

    def distorted_images(self):
        for Slide in self._Presentation.Slides:
            for Shape in Slide.Shapes:
                if self._Utils.is_image(Shape):
                    w, h = self._Utils.get_shape_percentage_width_height(Shape)
                    if abs(w - h) > 10:
                        return True
        return False

    def compare(self, path="original_images"):
        if Path(path).exists():
            return self.__compare_images(path=path)
        else:
            return "Не загружены изображения."

    def get(self, typeof="screenshot", lt="layout_1"):
        if typeof == "screenshot":
            return self.__generate_images_screenshots()
        elif typeof == "skeleton":
            return self.__generate_images_skeleton()
        elif typeof == "layout":
            return self.__generate_images_layout(lt)
        elif typeof == "thumb":
            return self.__generate_images_screenshots(thumb=True)
        else:
            return "Неожиданный тип генерации."
            # TODO: exception here

    @staticmethod
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
