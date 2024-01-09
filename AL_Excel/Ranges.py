""" AL_Excel.Ranges

    Class based structures for Cell Ranges
"""
## Super Module
from openpyxl import utils
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import Cell

## This Module
from AL_Excel import Coordinate, Index
from AL_Excel.Coordinates import CoordinateDescriptor

## Builtins
import typing

__all__ = ["NamedRange","Range"]

class NamedRange():
    def __init__(self,workbook:Workbook, name: str):
        self.workbook = workbook
        self.name = name
        if not workbook.defined_names[name]:
            raise ValueError(f'Workbook has no range named "{name}"')
        self.ranges = [Range(workbook[worksheet],_range) for worksheet,_range in workbook.defined_names[name].destinations]

    def allcells(self, sort: typing.Literal["row", "column"] = "row")-> typing.Generator[Cell,None,None]:
        """ Returns a flat list of all cells in the named range.

        if sort is "row" (default), the cells will be returned: first, in order
        of subrange; secondly, by row.
        if sort is "column", the second priority will be by column
        """
        if sort == "row": cells = [_range.rows_from_range("cell") for _range in self.ranges]
        elif sort == "column": cells = [_range.columns_from_range("cell") for _range in self.ranges]
        ## Range.[rows|columns]_from_range returns a generator for each range, which in turn returns 
        ## a generator to reference each [row|column] (so both need to be resolved via list and flattened)
        return sum([sum([list(cells) for cells in rowcol],[]) for rowcol in cells], [])

class Range():
    def __init__(self,worksheet: Worksheet, range: str):
        """ Creates a new Range Object

        worksheet is an openpyxl worksheet object
        range is a string expressing a valid, single range
        """
        self.worksheet = worksheet
        self._startcoord: Coordinate = None
        self._endcoord: Coordinate = None
        self._masterstart: Coordinate = None
        self._masterend: Coordinate = None
        coords = range.split(":")
        if 0 > len(coords) > 2:
            raise AttributeError(f"Invalid Range: {range}")
        self._startcoord = Coordinate(coords[0])
        if len(coords) == 2: self._endcoord = Coordinate(coords[1])
        else: self._endcoord = self.startcoord
        self.updatemasters()

    @property
    def startcoord(self)-> Coordinate:
        """ Returns the start coordinate of the Range"""
        return self._startcoord
    @startcoord.setter
    def startcoord(self,value: Coordinate|CoordinateDescriptor):
        """ Sets the start coordinate of the Range
        
            value can be a Coordinate Object or a CoordinateDescriptor
        """
        if not isinstance(value,Coordinate):
            value = Coordinate(value)
        self._startcoord = value
        self.updatemasters()
    @property
    def endcoord(self)-> Coordinate:
        """ Returns the end coordinate of the Range """
        return self._endcoord
    @endcoord.setter
    def endcoord(self,value: Coordinate|CoordinateDescriptor):
        """ Sets the end coordinate of the Range
        
            value can be a Coordinate Object or a CoordinateDescriptor
        """
        if not isinstance(value,Coordinate):
            value = Coordinate(value)
        self._endcoord = value
        self.updatemasters()

    @property
    def masterstart(self)-> Coordinate:
        """ Returns the start coordinate of the Range, relative to the worksheet """
        return self._masterstart
    @property
    def masterend(self)-> Coordinate:
        """ Returns the end coordinate of the Range, relative to the worksheet """
        return self._masterend
    
    def updatemasters(self):
        """ Updates the masterstart and masterend properties """
        c_1,r_1,c_2,r_2 = utils.cell.range_boundaries(self.range)
        self._masterstart = Coordinate(r_1,c_1)
        self._masterend = Coordinate(r_2,c_2)

    @property
    def range(self)-> str:
        """ Returns the range as an Excel-style string """
        return f"{self.startcoord.toA1string()}:{self.endcoord.toA1string()}"

    @property
    def range_boundaries(self) -> typing.Tuple[int, int, int, int]:
        """ Returns the results of openpyxl.utils.cell.range_boundaries, which is a tuple (min_col, min_row, max_col, max_row) """
        return utils.cell.range_boundaries(self.range)

    def rows_from_range(self,attribute: typing.Literal["value","address","cell"] = "value")-> typing.Generator[typing.Generator[Cell,None,None],None,None]:
        """ Returns cell property per openpyxl.utils.cell.rows_from_range
        
        If attribute is "address," functions as openpyxl.utils.cell.rows_from_range
        If attribute is "cell," returns cell references instead.
        """
        if attribute == "address":
            return utils.cell.rows_from_range(self.range)
        if attribute == "cell":
            return (map(lambda cell: self.worksheet[cell], row) for row in utils.cell.rows_from_range(self.range))
        return (map(lambda cell: getattr(self.worksheet[cell],attribute), row) for row in utils.cell.rows_from_range(self.range))

    def columns_from_range(self, attribute: typing.Literal["value","address","cell"]= "value")-> typing.Generator[typing.Generator[Cell,None,None],None,None]:
        """Returns cell property per openpyxl.utils.cell.rows_from_range
        
        If attribute is "address," functions as openpyxl.utils.cell.cols_from_range
        If attribute is "cell," returns cell references instead."""
        if attribute == "address":
            return utils.cell.cols_from_range(self.range)
        if attribute == "cell":
            return (map(lambda cell: self.worksheet[cell], column) for column in utils.cell.cols_from_range(self.range))
        return (map(lambda cell: getattr(self.worksheet[cell],attribute), column) for column in utils.cell.cols_from_range(self.range))

    def cells_by_row(self,attribute: typing.Literal["value","address","cell"] = "value")-> typing.Generator[Cell,None,None]:
        """ Returns a flat list of cells based on rows_from_range
        
        attribute is the same as rows_from_range
        """
        return (cell for row in list(self.rows_from_range(attribute=attribute)) for cell in row)

    def cells_by_column(self,attribute: typing.Literal["value","address","cell"] = "value")-> typing.Generator[Cell,None,None]:
        """ Returns a flat list of cells based on columns_from_range
        
        attribute is the same as rows_from_range
        """
        return (cell for column in list(self.columns_from_range(attribute=attribute)) for cell in column)

    def row(self,index,attribute: typing.Literal["value","address","cell"] = "value")-> typing.Generator[Cell,None,None]:
        """ Returns the row with the provided zero-index, via rows_from_range """
        return list(self.rows_from_range(attribute=attribute))[index]

    def column(self,index,attribute: typing.Literal["value","address","cell"] = "value")-> typing.Generator[Cell,None,None]:
        """ Returns the column with the provided zero-index, via columns_from_range """
        return list(self.columns_from_range(attribute=attribute))[index]

    def isinrange(self,cell: Cell)-> bool:
        """ Returns whether the cell is in the Range """
        coord = Coordinate(cell.coordinate)
        return (self.masterstart.column <= coord.column <= self.masterend.column)\
            and (self.masterstart.row <= coord.row <= self.masterend.row)

    def coordinatetocell(self, coordinate: Coordinate) -> Cell:
        """ Converts a Coordinate (which may be relative) to a Cell (which is absolute) based on the first index of the range """
        ## Note that range_boundaries is a CR pattern rather than RC
        if not coordinate.isabsolute():
            if not coordinate._row.absolute:
                if coordinate.row >= 0:
                    row  = self.masterstart.row + coordinate.row
                else:
                    row = self.masterend.row + coordinate.row + 1
                coordinate._row = Index("row",row,True)
            if not coordinate._column.absolute:
                if coordinate.column >= 0:
                    column = self.masterstart.column + coordinate.column
                else:
                    column = self.masterend.column + coordinate.column + 1
                coordinate._column = Index("column",column,True)
        return self.worksheet.cell(row = coordinate.row, column = coordinate.column)

    def subrange(self,start: Coordinate|CoordinateDescriptor = None, end: Coordinate|CoordinateDescriptor = None)-> "Range":
        """ Returns a new Range Object that is a slice of this Range Object

        Takes both start and end. start and end should be appropriate coordinates:
        an A1/R1C1 Notation string, tuple of (row,column) indexes, or Coordinate Object.
        Tuples follow the same syntax as Coordinate objects and must be length two.

        For tuples None can be substituted for either the value itself or either of the
        tuple indexes; as a replacement for start or end, it refers to the address of the
        (first-column,first-row) or the (last-column,last-row) respectively. As part
        of the tuple it refers to that part of the default value, i.e.- None for
        start.row is an alias for first-row, end.column = None references last-column, etc.

        Example - Range(C4:J6).subrange(end=(1,None)) == Range(C4:F4)

        String Notation and 
        """

        ## Note that range_boundaries is a CR pattern rather than RC
        if start is None: start = [(self.masterstart.row,True),(self.masterstart.column,True)]
        if isinstance(start,(tuple,list)):
            if len(start) != 2: raise ValueError(f"Start and end tuples must be length 2: {start}")
            start = list(start)
            if start[0] is None: start[0] = (self.masterstart.row,True)
            if start[1] is None: start[1] = (self.masterstart.column,True)
            start = Coordinate(*start)
        else: start = Coordinate(start)
        start = self.coordinatetocell(start)
        
        if end is None: end = [(self.masterend.row,True),(self.masterend.column,True)]
        if isinstance(end,(tuple,list)):
            if len(end) != 2: raise ValueError(f"Start and end tuples must be length 2: {end}")
            end = list(end)
            if end[0] is None: end[0] = (self.masterend.row,True)
            if end[1] is None: end[1] = (self.masterend.column,True)
            end = Coordinate(*end)
        else: end = Coordinate(end)
        end = self.coordinatetocell(end)

        
        if not self.isinrange(start) or not self.isinrange(end):
            start = Coordinate(start.row,start.column)
            end = Coordinate(end.row, end.column)
            raise ValueError(f"Subrange coordinates out of Range: R[{self.masterstart.row}]C[{self.masterstart.column}] : R[{self.masterend.row}]C[{self.masterend.column}] <> R[{start.row}]C[{start.column}] : R[{end.row}]C[{end.column}]")

        return Range(self.worksheet,f"{start.column_letter}{start.row}:{end.column_letter}{end.row}")

    def __eq__(self,other: "Range")-> bool:
        if isinstance(other,Range):
            return self.worksheet == other.worksheet and self.startcoord == other.startcoord and self.endcoord == other.endcoord
        
    def __add__(self, other: typing.Tuple[int|None, int|None]):
        return self.endcoord + other

    def __repr__(self):
        return f"{self.__class__.__name__}({self.range})"

def tuple_to_range(trange:typing.Tuple[int|str, int|str, int|str, int|str]|list[int|str, int|str, int|str, int|str], absolute:bool = False)-> str:
    """ Returns an Excel-version (A1 notation) string representing the given range.
        
        The tuple should be 1-indexed and should represent (startcolumn, startrow, endcolumn, endrow).

        The absolute keyword determines whether the returned range is absolute (default False).
    """
    c1,r1,c2,r2 = trange
    if absolute:
        c1,r1,c2,r2 = int(c1),int(r1),int(c2),int(r2)
    else:
        c1,r1,c2,r2 = str(c1),str(r1),str(c2),str(r2)
        
    start,end = Coordinate(r1,c1),Coordinate(r2,c2)
    return f"{start.toA1string()}:{end.toA1string()}"

def get_named_cell(workbook: Workbook, name:str)-> Cell:
    """ Resolves the given name as a NamedRange and if the NamedRange only has one Range
        with one Cell, it will return that cell.
        
        If the name does not exist, a ValueError will be raised.
        If the NamedRange exists but there is more than one range and/or more than one
        cell, an AttributeError will be raised.
    """

    nr = NamedRange(workbook,name)
    if len(nr.ranges) != 1:
        raise AttributeError("Defined Name does not refer to a single Cell.")
    cells = list(nr.allcells())
    if len(cells) != 1:
        raise AttributeError("Defined Name does not refer to a single Cell.")
    return cells[0]