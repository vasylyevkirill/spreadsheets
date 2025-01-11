import string
from .models import parse_coordinate, Coordinates

def test_parse_coordinate():
    coordinate, _ = parse_coordinate('12', '1234567890')
    assert coordinate == 12

    coordinate, _ = parse_coordinate('aa', string.ascii_lowercase)
    assert coordinate == 27

    coordinate_y, str_lost = parse_coordinate('A1', '1234567890')
    assert str_lost == 'A'
    coordinate_x, str_lost = parse_coordinate(str_lost.lower(), string.ascii_lowercase)
    assert coordinate_y == 1
    assert coordinate_x == 1
    assert str_lost == ''


def test_create_coordinate():
    coordinate = Coordinates('A1')
    assert coordinate.x == 0
    assert coordinate.y == 0
    assert coordinate.is_valid() == True

    coordinate = Coordinates()
    assert coordinate.is_valid() == False


