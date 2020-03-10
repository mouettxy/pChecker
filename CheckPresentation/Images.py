import os
import base64

from CheckPresentation import Data, Utils
from globals import DEFAULT_FONT

from PIL import Image, ImageDraw, ImageFont


class Images(Data.Data):
    """
    Класс реализовывает создание скриншотов презентации, вытаскивания из неё картинок, создания скелета слайдов
    презентации
    """
    def generate_skeleton(self, flag=None):
        """
        Генерирует и удаляет показательные изображения коллизий.
        :param flag: Список путей файлов которые необходимо удалить
        :type flag: list of string
        :return: список путей если флаг None, True в случае успешного удаления иначе
        :rtype: list of string or boolean
        """
        if flag is None:
            result = []
            for slide in self.presentation.slides:
                index = int(self.presentation.slides.index(slide) + 1)
                im = Image.new(mode="RGB", size=self.slide_dimensions())
                draw = ImageDraw.Draw(im)
                # previous_dimensions = []
                # previous_shape = None
                for shape in slide.shapes:
                    dim = self.shape_dimensions(shape)
                    if self.is_text(shape):
                        dim = self.shape_dimensions(shape)
                    # current_dimensions = [dim['x1'], dim['y1'], dim['width'], dim['height']]
                    image_color = "blue"
                    text_color = "yellow"
                    # if len(previous_dimensions) > 0 and previous_shape is not None:
                    # if self.check_collision(current_dimensions, previous_dimensions):
                    # print(self.check_collision(current_dimensions, previous_dimensions), shape.name,
                    # previous_shape.name)
                    if self.is_text(shape):
                        draw.rectangle([(dim['x1'], dim['y1']), (dim['x2'], dim['y2'])], fill=text_color, outline="red")
                    if self.is_image(shape):
                        draw.rectangle([(dim['x1'], dim['y1']), (dim['x2'], dim['y2'])], fill=image_color)
                    # previous_dimensions = current_dimensions
                    # previous_shape = shape
                im.save(f'{index}.jpg')
                result.append(f'{index}.jpg')
            return result
        else:
            for f in flag:
                os.mkdir(f)
            return True

    def save(self):
        """
        :return: Пути к сохранённым картинкам и их координаты Top Left на слайде
        :rtype: object of list of tuple (string, tuple of (int, int)
        """
        slide_counter, image_cords, image_paths = 0, [], []
        for slide in self.presentation.slides:
            slide_counter += 1
            picture_counter = 0
            for shape in slide.shapes:
                if self.is_image(shape):
                    picture_counter += 1
                    temp_image_paths = [
                        f"slide{slide_counter}_{picture_counter}.png",
                        f"slide{slide_counter}_{picture_counter}_original.png",
                    ]
                    image_size = (self.convert_emu_px(shape.width), self.convert_emu_px(shape.height))
                    original_image = open(temp_image_paths[1], 'wb')
                    original_image.write(base64.b64decode(base64.b64encode(shape.image.blob)))
                    original_image.close()
                    Image.open(temp_image_paths[1]).resize(image_size, Image.ANTIALIAS).save(temp_image_paths[0])
                    os.remove(temp_image_paths[1])
                    image_cords.append((self.convert_emu_px(shape.left), self.convert_emu_px(shape.top)))
                    image_paths.append(temp_image_paths[0])
        return list(zip(image_paths, image_cords))

    def generate(self, bg="white", text_fill="black"):
        """
        Генерирует максимально точно возможные скриншоты слайдов не используя API PowerPoint. Используя убедитесь что
        у пользователя есть доступ к папке C:\\Windows\\Fonts\\, и в том что установлены стандартные шрифты.
        Изображения генерируются в папке с файлом.
        :param bg: Желаемый цвет заднего фона
        :type bg: string
        :param text_fill: Желаемый цвет текста
        :return: Список с путями к сохранённым фотографиям
        :rtype: list of string
        """
        result = []
        images = self.save()
        prs_size = self.slide_dimensions()
        for slide in self.presentation.slides:
            idx = self.presentation.slides.index(slide) + 1
            slide_images = [(p[0], p[1]) for p in images if f'slide{idx}' in p[0]]  # p[0]: path ; p[1] width, height
            screen = Image.new("RGBA", prs_size, bg)
            for shape in slide.shapes:
                if self.is_text(shape):
                    dims = self.shape_dimensions(shape)
                    typeface = self.typefaces_by_shape(shape).split(' ')[0].lower()
                    font_size = int(max(self.font_sizes_by_shape(shape, True))) + 4
                    try:
                        font = ImageFont.truetype(font=f"C:/Windows/Fonts/{typeface}.ttf", size=font_size)
                    except OSError:
                        font = ImageFont.truetype(font=f"C:/Windows/Fonts/{DEFAULT_FONT.lower()}.ttf", size=font_size)
                    wrap = Utils.TextWrapper(shape.text, font, dims['width'])
                    text_image = Image.new("RGB", wrap.total_width_height(), bg)
                    draw = ImageDraw.Draw(text_image)
                    draw.text((0, 0), wrap.wrapped_text(), font=font, fill=text_fill)
                    screen.paste(text_image, (dims['x1'], dims['y1']))
                    # text_image.save(f'tmp/{self.random_string()}.jpg')
                    # process images
            for img in slide_images:
                screen.paste(Image.open(img[0]), img[1])
                os.remove(img[0])
            screen.save(f'slide{idx}.png')
            result.append(f'{os.getcwd()}\\slide{idx}.png')
        return result
