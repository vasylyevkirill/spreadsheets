import string
from .models import Spreadsheet, Coordinates


def test_spreadsheet_create():
    s = Spreadsheet(5, 5)
    c = Coordinates(x=1, y=1)
    assert s[c].value == None


def test_spreadsheet_shape(): 
    s = Spreadsheet(5, 5)
    assert s.shape[0] == 5
    assert s.shape[1] == 5
