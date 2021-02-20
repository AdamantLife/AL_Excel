## Test Target
from AL_Excel import Coordinates
## Test Framework
import unittest
## Testing utilities
from AL_Excel import tests
## Builtin
import builtins

class CoordinateTests(unittest.TestCase):
    def setUp(self):
        tests.basicsetup(self)
        return super().setUp()
    
    def test_creation(self):
        """ Test the creation of a Coordinate via various methods and validity of data """
        for coordinate in tests.DATA['COORDINATES']:
            with self.subTest(coordinate = coordinate):
                result = coordinate.get('result')
                if not result: continue
                definition = coordinate.get('definition')
                objecttype = coordinate.get("object")
                if objecttype:
                    factory = getobjectreference(objecttype)
                    definition = factory(definition)

                ## Tests
                if result['type'] == "success":
                    if isinstance(definition,(tuple,list)):
                        coord = Coordinates.Coordinate(*definition)
                    else:
                        coord= Coordinates.Coordinate(definition)
                    self.assertEqual(coord._row.value,result['row']['value'])
                    self.assertEqual(coord._row.absolute,result['row']['absolute'])
                    self.assertEqual(coord._column.value,result['column']['value'])
                    self.assertEqual(coord._column.absolute,result['column']['absolute'])
                elif result['type'] == "exception":
                    error = getattr(builtins,result['e_type'])

                    ## Added length check to handle different length objects
                    ## (that are expected to fail further downstream)
                    if isinstance(definition,(tuple,list)) and 0 < len(definition) < 3:
                        self.assertRaisesRegex(error,result['e_regex'],
                                               Coordinates.Coordinate,*definition)
                    elif isinstance(definition,dict):
                        self.assertRaisesRegex(error,result['e_regex'],
                                               Coordinates.Coordinate,**definition)
                    else:
                        self.assertRaisesRegex(error,result['e_regex'],
                                               Coordinates.Coordinate,definition)

    def test_coordinate_addition(self):
        """ Tests that Coordinate.__add__ works. """
        c1 = Coordinates.Coordinate(1,1)
        c2 = Coordinates.Coordinate("1","1")
        c3 = c1 + c2
        self.assertEqual(c3,Coordinates.Coordinate(2,2))

if __name__ == "__main__":
    unittest.main()