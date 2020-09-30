import pytest
from decorator import power, square

@pytest.mark.parametrize('a, b, c, answer', [
    (3, 3, 1, 8),
    (2, 2, 1, 1),
    (1, 1,'vasya', [])
])
def test_decoratorPower(a, b, c, answer):
    @power(pow=a)
    def f1(x, y):
        return x - y
    assert f3(b, c) == answer


@pytest.mark.parametrize('a, b, answer', [
    (3, 3, 36),
    (2, 2, 1),
    (1,9,[])
])
def test_decoratorSquare(a, b, answer):
    @square
    def f2(x, y):
        return x + y
    assert f2(a, b) == answer