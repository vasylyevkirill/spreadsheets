import string
from dataclasses import dataclass
from datetime import date
from pandas import DataFrame


def parse_coordinate(str_coordinate: str, expectex_symbols: list):
    coordinate: int = 0
    index: int = 0
    for index, symbol in enumerate(reversed(str_coordinate)):
        if symbol in expectex_symbols:
            coordinate += (expectex_symbols.index(symbol) + 1) * (len(expectex_symbols) ** index)
            index += 1
            continue
        return coordinate, str_coordinate[:-(index)]
    return coordinate, ''


class Expression():
    def __init__(self, str_expression: str):
        self._value = str_expression


@dataclass()
class Coordinates():
    EXPECTEX_X_SYMBOLS = string.ascii_lowercase
    EXPECTEX_Y_SYMBOLS = '1234567890'
    x: int
    y: int

    def __init__(self, coordinate: str | None = None, x: int = -1, y: int = -1):
        if not coordinate:
            self.x = x
            self.y = y
        elif coordinate:
            self.parse_str_coordinate(coordinate)
        else:
            raise ValueError('Coordinate not passed')

    def parse_str_coordinate(self, str_coordinate: str) -> (int, int):
        str_coordinate = str_coordinate.lower().strip()
        self.y, str_coordinate = parse_coordinate(str_coordinate, self.EXPECTEX_Y_SYMBOLS)
        self.x, str_coordinate = parse_coordinate(str_coordinate, self.EXPECTEX_X_SYMBOLS)

        self.x -= 1
        self.y -= 1

        return self.x, self.y
            
    def is_valid(self, raise_exception: bool = False) -> bool:
        if self.x >= 0 and self.y >= 0:
            return True
        return False

@dataclass()
class SpreadsheetCell():
    coordinates: Coordinates
    value: str | int | date | Expression | None

    def calculate_value(self, data_format: str | None):
        if isinstance(self.value, Expression):
            return self.value.calculate_value()
        return self.value

    @property
    def display_value(self) -> str | int | date | None:
        return self.calculate_value()


class Spreadsheet():
    def __init__(self, len_x: int = -1, len_y: int = -1, data: dict | list | None = None):
        if not data:
            data = [[None for i in range(len_x)] for i in range(len_y)]
        self._dataframe = DataFrame(data=data)

    def __getitem__(self, coordinates: Coordinates) -> SpreadsheetCell:
        _dataframe[coordinates]

 

class SpreadsheetExpression():
    def __init__(self, expression):
        self._expression = expression

    def calculate_value(self):
        return
