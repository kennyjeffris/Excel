# Importable functions
from string import ascii_letters

def as_text(value):
    """Return the input as a string."""
    if value is None:
        return ""
    return str(value)


def poly_fit(x, coefficients):
    return (coefficients[3] + (coefficients[0] - coefficients[3]) /
            ((1 + (x / coefficients[2]) ** coefficients[1]) ** coefficients[4]))


def col2num(colindex):
    """Convert excel column letter to index number."""
    number = 0
    for c in colindex:
        if c in ascii_letters:
            number = number * 26 + (ord(c.upper()) - ord('A')) + 1
    return number