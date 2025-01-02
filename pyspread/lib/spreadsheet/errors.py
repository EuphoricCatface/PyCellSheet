class SpreadsheetErrorBase(Exception):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "BASE?!"

    def cell_output(self):
        return f"#{self.__string}"


class SpreadsheetErrorNull(SpreadsheetErrorBase):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "NULL!"


class SpreadsheetErrorDivZero(SpreadsheetErrorBase):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "DIV/0!"


class SpreadsheetErrorValue(SpreadsheetErrorBase):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "VALUE!"


class SpreadsheetErrorRef(SpreadsheetErrorBase):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "REF!"


class SpreadsheetErrorName(SpreadsheetErrorBase):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "NAME?"


class SpreadsheetErrorNum(SpreadsheetErrorBase):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "NUM!"


class SpreadsheetErrorNa(SpreadsheetErrorBase):
    def __init__(self, *args):
        super().__init__(*args)
        self.__string = "N/A"
