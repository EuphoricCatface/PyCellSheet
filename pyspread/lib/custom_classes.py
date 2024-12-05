# Created by EuphoricCatface


EmptyCell = object()


class PythonCode(str):
    pass


class Range(list):
    def __init__(self):
        super().__init__()
        self.top_left = None
