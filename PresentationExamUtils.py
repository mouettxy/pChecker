from MSOCONSTANTS import msoTrue, msoPicture, msoLinkedPicture, msoAutoShape, msoPlaceholder
from MSOCONSTANTS import ppPlaceholderPicture


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
        if Shape.Type == msoPicture or Shape.Type == msoLinkedPicture or Shape.Type == msoAutoShape:
            return True
        if Shape.Type == msoPlaceholder:
            if Shape.PlaceholderFormat.Type == ppPlaceholderPicture:
                return True
        return False

    @staticmethod
    def check_collision_between_shapes(first_shape, second_shape):
        if (first_shape['left'] + first_shape['width'] >= second_shape['left'] and
                first_shape['left'] <= second_shape['left'] + second_shape['width'] and
                first_shape['top'] + first_shape['height'] >= second_shape['top'] and
                first_shape['top'] <= second_shape['top'] + second_shape['height']):
            return True
        return False

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
                    shape_t = self.convert_points_px(Range.BoundTop) - self.convert_points_px(TextFrame.MarginTop)
                    shape_l = self.convert_points_px(Range.BoundLeft) - self.convert_points_px(TextFrame.MarginLeft)
                    shape_w = self.convert_points_px(Range.BoundWidth) - self.convert_points_px(TextFrame.MarginRight)
                    shape_h = self.convert_points_px(Range.BoundHeight) - self.convert_points_px(TextFrame.MarginBottom)
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