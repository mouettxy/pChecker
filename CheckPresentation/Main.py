from CheckPresentation import Analyze, Images, Warnings, Testing, Data, Utils


class Main(object):

    def __init__(self, path_to_presentation,):
        super().__init__()
        self.Analyze = Analyze.Analyze(path_to_presentation)
        self.Data = Data.Data(path_to_presentation)
        self.Images = Images.Images(path_to_presentation)
        self.Testing = Testing.Testing(path_to_presentation)
        self.Utils = Utils.Utils(path_to_presentation)
        self.Warnings = Warnings.Warnings(path_to_presentation)
