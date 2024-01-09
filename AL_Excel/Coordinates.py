""" AL_Excel.Coordinates

    Implementation of a Coordinate Class which makes it easier to handle locations
"""
## Super Module
from openpyxl import utils
## Builtin
import re
import typing


__all__ = ["Coordinate", "Index"]

IndexType = typing.Literal["row","column"]

class Index(typing.NamedTuple):
    """ Represents a single row or column index

    Attributes:
        type: "row" or "column"
        value: the index value
        absolute: True or False
    """
    type: IndexType
    value: int|None
    absolute: bool

## These can be parsed into an Index
IndexDescriptorTuple = typing.Tuple[int|str|Index,typing.Optional[bool]]
IndexDescriptorList = list[int|str|Index, typing.Optional[bool]]
IndexDescriptor = None|str|int|Index|IndexDescriptorTuple|IndexDescriptorList

CoordinateTupleDescriptor = typing.Tuple[IndexDescriptor,IndexDescriptor]
CoordinateListDescriptor = list[IndexDescriptor,IndexDescriptor]
CoordinateDescriptor = str|CoordinateTupleDescriptor|CoordinateListDescriptor

class Coordinate():
    """ Represents a row and/or column index

    When both row and column are supplied, represents a specific cell index (e.g.- [Cell] A1).
    If the the first argument is a string and no other arguments are passed, the string should represent a Excel-format
    coordinate (either A1 or R1C1).
    If one of row or column is supplied and the other is None, represents a row or column index, respectively (i.e.- [Column] A).

    Row or column may individually be defined as a Excel-formatted string, an integer, or a tuple. Positive integers are always
    treated as absolute references where as strings are only treated as relative by default. A tuple should contain either
    the integer or string index and can have an optional second index indicating that it is an absolute reference (supplying
    either True, "$", or "absolute"). Row and column indexes can be negative only so long as they are not absolute (negative integers
    will be treated as relative).

    Each part of a Coordinate instance (row and column) is a namedtuple called Index with parts (type, value, absolute) where type indicates "row" or "column",
    value is the reference, and absolute is a boolean (True or False).
    """
    def __init__(self,row: IndexDescriptor|CoordinateDescriptor|None, column: IndexDescriptor|None = False):
        if isinstance(row,str) and column is False:
            row,column = converttotuple(row)
        else:
            if column is False: column = None
            try: row,column = converttotuple((row,column))
            except Exception as e:
                raise AttributeError(f"Invalid values for cooridinate: {row},{column}")
        self._row = row
        self._column = column
    @property
    def row(self)->int|None:
        return self._row.value if self._row else None
    @property
    def column(self)->int|None:
        return self._column.value if self._column else None
    def isabsolute(self)->bool:
        return self._column.absolute or self._row.absolute
    def toA1string(self, absolute: bool = True)->str:
        """ Returns the Coordinate in A1 format.

            By default, will output absolute dollar signs (ex.- A$1).
            If absolute is False, will output both row and column as relative.
        """
        if not absolute:
            return f"{utils.cell.get_column_letter(self._column.value)}{self._row.value}"
        return f'{"$" if self._row.absolute else ""}{utils.cell.get_column_letter(self._column.value)}{"$" if self._column.absolute else ""}{self._row.value}'
    
    def __add__(self,other: "Coordinate"|IndexDescriptor|CoordinateDescriptor|None)->"Coordinate":

        if isinstance(other,Coordinate):
            if (self._row.absolute and other._row.absolute)\
               or self._column.absolute and other._column.absolute:
                raise AttributeError(f"Cannot add two Absolute references: {self} + {other}")
            
            return Coordinate(row = addindices(self._row,other._row), column = addindices(self._column, other._column))
        
        ## Otherwise, try to convert to Coordinate
        ## Conversion will automatically handle Address Strings and 2-length iterables (i.e. (row,column) Tuple)
        try: return self + Coordinate(other)
        except: pass

    def __eq__(self,other: "Coordinate")->bool:
        if isinstance(other,Coordinate):
            return self._row == other._row and self._column == other._column
        
    def __repr__(self):
        return f"{self.__class__}({self._row},{self._column})"



## Order than Coordinates should be displayed
COORDINATEORDER = ["row","column"]

FULLRE = re.compile("""
^(?P<match>
    (?P<RC>
        (?P<RCrow>R                                                 ## R1C1
            (?P<absolute1>\[?)(?P<RCrowid>-?\d+)(?P<absolute2>\]?)
        )
        (?P<RCcolumn>C
            (?P<absolute3>\[)?(?P<RCcolumnid>-?\d+)(?P<absolute4>\]?)
        )
    )
    |
    (?P<A1>                                                         ## A1
        (?P<A1column>
            (?P<absoluteB>\$?)(?P<A1columnid>[A-Z]+)
        )
        (?P<A1row>
            (?P<absoluteA>\$?)(?P<A1rowid>\d+)
        )
    )                 
)$
""", re.VERBOSE | re.IGNORECASE)
COLUMNRE = re.compile("""
(?P<match>
      C(?P<absolute1>\[?)(?P<RCcolumn>-?\d+)(?P<absolute2>\]?)    ## R1C1
    | (?P<absolute>\$?)(?P<A1column>[A-Z]+)                     ## A1
)
""", re.VERBOSE | re.IGNORECASE)
ROWRE = re.compile("""
(?P<match>
      R(?P<absolute1>\[?)(?P<RCrow>-?\d+)(?P<absolute2>\]?)   ## R1C1
    | (?P<absolute>\$?)(?P<A1row>\d+)                       ## A1
)
""", re.VERBOSE | re.IGNORECASE)

def converttotuple(value: CoordinateDescriptor) -> tuple[Index,Index]:
    """ Converts a value into a tuple of Indexes ambiguously representing a row and column.

        Acceptable values:
            * A string representing RC or A1 Notation for a Coordinate/Cell
            * A tuple or list representing a row and column index which is a valid value for parseindex

        Parameters:
            value: The value to parse

        Returns:
            tuple[Index,Index]: A tuple of Indexes representing a row and column index
    """
    if isinstance(value,str):
        regex = FULLRE.search(value)
        if regex:
            syntax = "RC" if regex.group("RC") else "A1"
            row = parsecoordregex(regex.group(f"{syntax}row"),'row')
            column = parsecoordregex(regex.group(f"{syntax}column"),'column')
            return row,column
        else:
            raise TypeError(f"Invalid coordinates: {value}")
        
    
    try: index1,index2 = value
    except: raise TypeError(f"Cooridinates must be Coordinate-Formatted string or a tuple: {value}")

    indices = [parseindex(index1),parseindex(index2)]

    nonindices = [index.value for index in indices if index.value is None]

    ## Both are None indexes
    if len(nonindices) >= 2:
        raise ValueError(f"No Indicies provided: {indices}")
    
    if nonindices:
        goodind = 0 if indices[0].value is not None else 1
        goodindex = indices[goodind]

        ## If goodindex is ambiguous, then we need to set it's type
        if goodindex.type is None:
            ## type is determined by position in the tuple 
            ## i.e. (None,1) -> (NoneIndex,AmbiguousIndex==1) -> (NoneRowIndex,ColumnIndex==1)
            goodindex = goodindex._replace(type = COORDINATEORDER[goodind])

        ## Otherwise, we need to double-check goodindex's position in the tuple
        else:
            if goodindex.type != COORDINATEORDER[goodind]:
                indices[0],indices[1] = indices[1],indices[0]

        ## Output will be (NoneRow, column) or (row, NoneColumn)
        return indices

    ## row/column are ambiguous
    if indices[0].type is None and indices[1].type is None:
        ## Assign their types in order
        indices[0] = indices[0]._replace(type = "row")
        indices[1] = indices[1]._replace(type = "column")
        return indices

    ## Indices are of the same type: the only way this should happen is because of the User
    ## e.g.- Coordinate("A","A"), Coordinate("1","1")
    if indices[0].type == indices[1].type:
        ## If both are rows, then we can try assume one to actually be a column index
        if indices[0].type == "row" and indices[1].type == "row":
            ## If only one is absolute, then the other can be assumed to be relative
            if indices[0].absolute and not indices[1].absolute:
                indices[1] = indices[1]._replace(type='column')
            elif indices[1].absolute and not indices[0].absolute:
                indices[1] = indices[0]._replace(type='column')
            elif not indices[0].absolute and not indices[1].absolute:
                indices[1] = indices[1]._replace(type='column')
            else:
               raise ValueError(f"Duplicated Arguments: {indices[0].type},{indices[1].type}")
        else:
            raise ValueError(f"Duplicated Arguments: {indices[0].type},{indices[1].type}")

    row = [index for index in indices if index.type == "row"]
    column = [index for index in indices if index.type == "column"]
    ambiguous = [index for index in indices if index.type is None]

    if row and not column:
        row = row[0]
        column = ambiguous[0]
        column = column._replace(type = "column")
    elif column and not row:
        column = column[0]
        row = ambiguous[0]
        row = row._replace(type = "row")
    else:
        row = row[0]
        column = column[0]

    return row,column

    

def parseindex(value: IndexDescriptor)->Index:
    """ Parses a value into an Index. If it is already an Index, confirms that it is valid.
    
        Valid Index Formats:
            * A string representing RC or A1 Notation
            * An integer (as a row or column index)
            * A tuple or list with a single value (row or column index) or two values (row or column index, absolute)
            * An Index namedtuple

        Parameters:
            value: The value to parse

        Returns:
            Index: A namedtuple with parts (type, value, absolute)
    """
    ## TODO: Determine if a None-Index is a valid use-case
    if value is None:
        return Index(None,None,False)
    
    if isinstance(value,int) and not isinstance(value,bool):
        return Index(None,value,True)
    
    if isinstance(value,str):
        try:
            return parsecoordregex(value,"column")
        except:
            return parsecoordregex(value,"row")
        
    ## Validate a Index
    ## Note, namedtuple is an instance of tuple, so this has to happen before handling tuple/lists
    if isinstance(value,Index):
        ## Index should be "row","column" (, or None for Column References)
        if value.type not in ["row","column",None]: raise ValueError(f"Index's type should be 'row','column', or None: {value}")
        ## Index made by this module use column index (or None for Column References)
        if not isinstance(value.value,int) and value.value is not None: raise ValueError(f"Index's value is not an integer (or None): {value}")
        if value.absolute not in [True,False]: raise ValueError(f"Index's absolute value is not True or False: {value}")
        return value

    if isinstance(value,(tuple,list)):
        if 0 > len(value) or len(value) > 2:
            raise ValueError(f"Coordinate part's length should be 1 or 2 (index,[absolute]): received length {len(value)}")
        
        ## Recurse the first value
        index = parseindex(value[0])

        if len(value) == 2:
            absolute = value[1]
            if isinstance(absolute,str):
                absolute = absolute.lower()
                ## Acceptable absolute = True values
                if absolute in ["$","absolute"]: absolute = True
                ## Only absolute = False value
                elif absolute == "": absolute = False
                else: raise ValueError(f"Coordinate part's second index is an unknown string: {value[1]}")

            if absolute not in [0,1,True,False]:
                raise ValueError(f"Coordinate part's second index must be True or False, or an accepted alias: {absolute}")
            
            ## Check collision between index defined in value[0]:
            ## ("$A",False) seems to be the only relevant case
            if index.absolute and not absolute:
                raise ValueError(f"Coordinate part's second index contradicts the first: {value}")
            
            index = Index(index.type, index.value, absolute)
        return index

    ## Everything else
    raise ValueError(f"Coordinate part is not a recognized format: {value}")

def parsecoordregex(value: str,colrow: IndexType)->Index:
    """ Parses a string using a regex to determine the index and absolute value
    
    Parameters:
        value: The string to parse
        colrow: Whether the string is a column or row index
    
    Returns:
        Index: A namedtuple with parts (type, value, absolute)
    """
    if colrow not in ("column","row"): raise ValueError(f"parsecoordregex must be 'column' or 'row': {colrow}")

    if colrow == "column": regex = COLUMNRE.search(value)
    else: regex = ROWRE.search(value)
    if not regex: raise ValueError(f"String does not match any identifiable {colrow} patterns: {value}")

    if regex.group(f"A1{colrow}"):
        index = regex.group(f"A1{colrow}")
        if colrow == "column": index = utils.cell.column_index_from_string(index)
        index = int(index)
        return Index(colrow, index, bool(regex.group("absolute")))
    else:
        abs1,abs2 = bool(regex.group("absolute1")),bool(regex.group("absolute2"))
        ## If both are not true, then we are missing one of the brackets
        if abs1 != abs2:
            ## If abs1 is True, then we are missing the closing bracket (abs2)
            missing = "]" if abs1 else "["
            raise SyntaxError(f"Incomplete coordinate description: missing {missing}.")
        
        return Index(colrow, int(regex.group(f"RC{colrow}")),abs1 and abs2)

def addindices(index1: Index,index2: Index)-> Index:
    """ Adds two indices together. Both must have the same type and at least one must be relative (Index.absolute == False)
    
    Parameters:
        index1: The first index
        index2: The second index
    
    Returns:
        Index: The sum of the two indices
    """
    if index1.absolute and index2.absolute:
        raise AttributeError(f"Cannot add two Absolute Indices: {index1} + {index2}")
    
    if index1.type != index2.type:
        raise TypeError(f"Type mismatch between Indices: {index1}, {index2}")
    
    ## Sort indices so that any absolute index is first
    ## (so only need to check if the second needs to be moved)
    if index2.absolute: index1,index2 = index2,index1

    ## If index1 is absolute, then the result should be absolute
    result = Index(index1.type, index1.value + index2.value, index1.absolute)

    ## If result is absolute, make sure that it is is still above 0
    ## (doesn't matter for relative Indexes)
    ## There doesn't seem to be any valid reason to use min(1), since the
    ## first index is always a known quantity and any user who wants to reduce
    ## a Coordinate/Index to the first index should just do so manually.
    if result.absolute and result.value <= 0:
        raise ValueError(f"Absolute Reference reduced below 0: {result}")
    return result