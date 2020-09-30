import pytest
from calc import line
from calc import cube
from calc import square

@pytest.mark.parametrize('a, b, answer', [
    (1, 'vanya', []),
    (2, 2, [5]),
    (3, 4, [-2])
])
def test_line(a, b, answer):
    assert line(a, b) == answer


@pytest.mark.parametrize('a, b, c, answer', [
    (1, 2, 1, [-1]),
    (2, 1, 3, [4]),
    (1, 'petya', 5, [])
])
def test_square(a, b, c, answer):
    assert square(a, b, c) == answer


@pytest.mark.parametrize('a, b, c, d, answer', [
    (1, 1, 1, 1, [-0.9999999999999998, (-1.1102230246251565e-16+1j), (-1.1102230246251565e-16-1j)]),
    (2, 4, 5, 1, []),
    (3, 1, 9, 12, [12])
])
def test_cube(a, b, c, d, answer):
    assert cube(a, b, c, d) == answer


def test_lineNegative(human):
    assert line() == []
