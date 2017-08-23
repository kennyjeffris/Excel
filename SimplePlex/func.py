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

def prep_lists(analytes):
    # Define lists as global
    searchList = []
    headerList = []
    sample_name_options = ['Sample', 'SampleName']
    searchList.append([sample_name_options, 'Gnr1Background', 'Gnr1RFU', 'Gnr2RFU', 'Gnr3RFU',
                     'Signal', 'RFUPercentCV', 'Gnr1Signal', 'Gnr2Signal', 'Gnr3Signal',
                     'RFU', 'Gnr1CalculatedConcentration',
                     'Gnr2CalculatedConcentration',
                     'Gnr3CalculatedConcentration', 'CalculatedConcentration',
                     'CalculatedConcentrationPercentCV'])

    headerList.append(['Sample #', 'Sample Name', 'Bkgd', 'Gnr1', 'Gnr2', 'Gnr3', 'Avg', '% CV', 'Gnr1', 'Gnr2',
                     'Gnr3', 'Avg', 'Gnr1', 'Gnr2', 'Gnr3', 'Avg', '% CV'])

    # Lists for Summary 2
    searchList.append([sample_name_options, 'Gnr1CalculatedConcentration',
                     'Gnr2CalculatedConcentration', 'Gnr3CalculatedConcentration', 'CalculatedConcentration',
                     'CalculatedConcentrationPercentCV'])

    headerList.append(['Sample #', 'Sample Name', 'Gnr1', 'Gnr2', 'Gnr3', 'Avg', '% CV'])

    # Lists for Summary 3
    headerList.append(analytes[:])
    headerList[2].insert(0, 'Sample')

    headerList.append(['CalculatedConcentration', 'CurveCoefficientA',
                     'CurveCoefficientB', 'CurveCoefficientC', 'CurveCoefficientD',
                     'CurveCoefficientG'])

    headerList.append(['Curve Coefficients', 'A', 'B', 'C', 'D', 'G'])
    return headerList, searchList