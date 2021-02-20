## Test Target
from AL_Excel import Ranges
## Test Framework
import unittest
## Testing utilities
from AL_Excel import tests
## This Module
import AL_Excel
## Builtin
import types

class RangeTests(unittest.TestCase):
    def setUp(self):
        tests.basicsetup(self)
        return super().setUp()

    def test_range__eq__(self):
        """ Tests various equivalencies between Range Objects"""
        wb = self.workbook
        cs = self.sheet_CellNumbers
        testrange = Ranges.Range(cs,"$A$1:$J$1")
        testrange2 = Ranges.Range(cs,"$A$1:$J$1")
        testrange3 = Ranges.NamedRange(wb,"testrange_1").ranges[0]
        self.assertEqual(testrange,testrange2)
        self.assertEqual(testrange,testrange3)

    def test_testranges(self):
        """ Make sure testranges line up with what is actually in the Excel file """
        for solutionname,solution in tests.DATA['RANGES'].items():
            with self.subTest(solutionname,solution = solution):
                testrange = Ranges.NamedRange(self.workbook,solution['name']).ranges[0]

                self.assertEqual([list(row) for row in testrange.rows_from_range("address")],solution['row_values'])
                self.assertEqual([list(row) for row in testrange.rows_from_range("value")],solution['row_values'])
                self.assertEqual([list(row) for row in testrange.rows_from_range("cell")],[list(self.sheet_CellNumbers[cell] for cell in row) for row in solution['row_values']])
        
    def test_rows_from_range(self):
        for solutionname,solution in tests.DATA['RANGES'].items():
            with self.subTest(solutionname=solutionname, solution = solution):
                testrange = Ranges.NamedRange(self.workbook,solution['name']).ranges[0]
                rows = testrange.rows_from_range("cell")
                ## Assert Generator
                self.assertIsInstance(rows,types.GeneratorType)
                rows = list(rows)
                ## Assert Number of Rows
                self.assertEqual(len(rows),len(solution['row_values']))
                ## Assert length of row
                ## Note: casting is due to generators and mapping being used
                self.assertEqual(len(list(list(rows)[0])),
                                 len(solution['row_values'][0]))
                ## Assert Value is Default
                self.assertEqual(list(list(row) for row in testrange.rows_from_range()),list(list(row) for row in testrange.rows_from_range("value")))
                ## Assert Addresses
                self.assertEqual(list(list(row) for row in testrange.rows_from_range("address")), solution['row_values'])
                ## Assert Values
                self.assertEqual(list(list(row) for row in testrange.rows_from_range("value")),solution['row_values'])
                ## Assert Cells-type
                self.assertTrue(all(
                    all(isinstance(cell,AL_Excel.cell.Cell) for cell in row) for row in rows
                    ))

class MethodsTest(unittest.TestCase):
    """ Tests for various utility methods of the Ranges module """
    def setUp(self):
        tests.basicsetup(self)

    def test_get_named_cell(self):
        """ Tests the get_named_cell utility function """
        cell = Ranges.get_named_cell(self.workbook, "TheFirstCell")
        self.assertIsInstance(cell, AL_Excel.cell.Cell)
        self.assertEqual(cell.coordinate, "A1")
        self.assertEqual(cell.value, "A1")

if __name__ == "__main__":
    unittest.main()