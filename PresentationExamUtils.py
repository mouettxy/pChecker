from MSOCONSTANTS import msoTrue, msoPicture, msoLinkedPicture, msoPlaceholder
from MSOCONSTANTS import ppPlaceholderCenterTitle, ppPlaceholderTitle, ppPlaceholderSubtitle
from MSOCONSTANTS import ppPlaceholderPicture, msoScaleFromTopLeft


class PresentationExamUtils(object):
    def __init__(self, application):
        super().__init__()
        self._Application = application

    @staticmethod
    def convert_points_px(value):
        return round(value / 72 * 96)

    @staticmethod
    def is_text(Shape):
        if Shape.HasTextFrame and Shape.Visible == msoTrue:
            if Shape.TextFrame.HasText:
                return True
            else:
                return None
        return False

    @staticmethod
    def is_image(Shape):
        if Shape.Type == msoPicture or Shape.Type == msoLinkedPicture:
            return True
        if Shape.Type == msoPlaceholder:
            if Shape.PlaceholderFormat.Type == ppPlaceholderPicture:
                return True
        return False

    @staticmethod
    def is_title(Shape):
        if Shape.Type == msoPlaceholder:
            if (Shape.PlaceholderFormat.Type == ppPlaceholderCenterTitle or
                    Shape.PlaceholderFormat.Type == ppPlaceholderSubtitle or
                    Shape.PlaceholderFormat.Type == ppPlaceholderTitle):
                return True
        return False

    @staticmethod
    def check_collision_between_shapes(first_shape, second_shape):
        if (first_shape['left'] + first_shape['width'] > second_shape['left'] and
                first_shape['left'] < second_shape['left'] + second_shape['width'] and
                first_shape['top'] + first_shape['height'] > second_shape['top'] and
                first_shape['top'] < second_shape['top'] + second_shape['height']):
            return True
        return False

    @staticmethod
    def dict_to_list(dictionary, key=None):
        if key is None:
            for d in dictionary:
                if type(dictionary[d]) == dict:
                    for d2 in dictionary[d]:
                        yield dictionary[d][d2]
                else:
                    yield dictionary[d]
        else:
            for d in dictionary[key]:
                yield dictionary[key][d]

    def get_shape_dimensions(self, Shape):
        if Shape.HasTextFrame:
            if Shape.TextFrame.HasText:
                movement_counter = 0
                text_length = Shape.TextFrame.TextRange.Length
                # if shape has text, we delete all end characters like enter, vertical tab, spaces, etc
                while (Shape.TextFrame.TextRange.Characters(text_length, 1).Text == chr(13) or
                       Shape.TextFrame.TextRange.Characters(text_length, 1).Text == chr(11) or
                       Shape.TextFrame.TextRange.Characters(text_length, 1).Text == chr(10) or
                       Shape.TextFrame.TextRange.Characters(text_length, 1).Text == chr(32)):
                    Shape.TextFrame.TextRange.Characters(text_length, 1).Delete()
                    text_length -= 1
                    movement_counter += 1
                else:
                    Range = Shape.TextFrame.TextRange
                    TextFrame = Shape.TextFrame
                    # int 7 that is experimental average value
                    shape_t = self.convert_points_px(Range.BoundTop) - self.convert_points_px(TextFrame.MarginTop) + 7
                    shape_l = self.convert_points_px(Range.BoundLeft) - self.convert_points_px(TextFrame.MarginLeft) + 7
                    shape_w = self.convert_points_px(Range.BoundWidth) - self.convert_points_px(
                        TextFrame.MarginRight) - 7
                    shape_h = self.convert_points_px(Range.BoundHeight) - self.convert_points_px(
                        TextFrame.MarginBottom) - 7
                    # undo all what we deleted, because for some reason, this changes saved in presentation
                    for i in range(movement_counter):
                        self._Application.StartNewUndoEntry()
                    return {
                        'left': shape_l,
                        'top': shape_t,
                        'width': shape_w,
                        'height': shape_h
                    }
        return {
            'top': self.convert_points_px(Shape.Top),
            'left': self.convert_points_px(Shape.Left),
            'width': self.convert_points_px(Shape.Width),
            'height': self.convert_points_px(Shape.Height),
        }

    def get_shape_crop_values(self, Shape):
        if not self.is_text(Shape):
            return {
                'left': self.convert_points_px(Shape.PictureFormat.CropLeft),
                'top': self.convert_points_px(Shape.PictureFormat.CropTop),
                'right': self.convert_points_px(Shape.PictureFormat.CropRight),
                'bottom': self.convert_points_px(Shape.PictureFormat.CropBottom),
            }
        else:
            return {'left': 0, 'top': 0, 'right': 0, 'bottom': 0}

    @staticmethod
    def get_shape_percentage_width_height(Shape, original_w_h=False):
        shape_width, shape_height = Shape.Width, Shape.Height
        Shape.ScaleWidth(1, msoTrue, msoScaleFromTopLeft)
        Shape.ScaleHeight(1, msoTrue, msoScaleFromTopLeft)
        original_width, original_height = (Shape.Width,
                                           Shape.Height)
        percentage_width, percentage_height = (shape_width / original_width * 100,
                                               shape_height / original_height * 100)
        Shape.ScaleWidth(percentage_width / 100, msoTrue)
        Shape.ScaleHeight(percentage_height / 100, msoTrue)
        if original_w_h:
            return round(original_width), round(original_height)
        return round(percentage_width), round(percentage_height)
