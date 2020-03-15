class PresentationExamLayouts(object):
    """
    Contains information about possible layouts of the 2nd and 3rd presentation slides.
    Each method is responsible for a particular layout. Each method gives a dictionary containing lists containing
    tuple of width and height that contains coordinates in px rectangles in which pictures, text blocks should be
    located.
    """

    @staticmethod
    def layout_1(width, height):
        """
        second slide: 2 images, 2 text blocks
        third slide: 3 images, 3 text blocks
        :param width: width of slide in pixels
        :type width: int
        :param height: height of slide in pixels
        :type height: int
        :return: dict of int: dict of string: list of tuple of (int, int)
        """
        return {
            2: {
                'images': [(0, 0, width/2, height)],
                'text_blocks': [(width/2, 0, width/2, height), (0, 0, width, height/4)]
            },
            3: {
                'images': [(0, 0, width/3, height/2), (width/3, height/2, width/3+width/3, height/2),
                           (width/3*2, 0, width/3, height/2)],
                'text_blocks': [(0, height/2, width/3, height/2), (width/3, 0, width/3, height/2),
                                (width/3*2, height/2, width/3, height/2), (0, 0, width, height/4)],
            }
        }
