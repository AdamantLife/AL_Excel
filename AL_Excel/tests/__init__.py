## Tested Module
import AL_Excel
## Builtin
import json
import pathlib
import unittest

DIRECTORY = pathlib.Path(__file__).resolve().parent
DATAFILE = (DIRECTORY / "exceltests.json").resolve()
TESTFILE = (DIRECTORY / "testbook.xlsx").resolve()

def loaddata():
    """ Loads testdata from DATAFILE """
    with open(DATAFILE,'r') as f:
        data = json.load(f)
    ## Time-saving for tests, space-saving for jsonfile, avoids clerical errors
    for _range in data['RANGES'].values():
        _range['column_values'] = list(zip(*_range['row_values']))
    return data

DATA = loaddata()

def getobjectreference(objecttype):
    """ Searches the global scope and the excel module for the provided reference string """
    try: factory = globals[objecttype]
    except KeyError:
        factory = getattr("AL_Excel",objecttype)
        if not factory: raise ValueError(f"Test could not build: {objecttype}")
    return factory

def lambdafactory(*args,**kw):
    """ A callable to create a lambda for testing """
    return lambda *args, **kwargs: None

class TestObject():
    """ A simple object for testing """
    pass

def basicsetup(testcase):
    """ Does basic workbook setup for the testcase """
    testcase.workbook = AL_Excel.load_workbook(str(TESTFILE), data_only = True)
    for sheet in DATA['SHEETS']:
        setattr(testcase,sheet['alias'],testcase.workbook[sheet['name']])


if __name__ == "__main__":
    import unittest
    import pathlib
    path = pathlib.Path.cwd()
    tests = unittest.TestLoader().discover(path)
    #print(tests)
    unittest.TextTestRunner().run(tests)
