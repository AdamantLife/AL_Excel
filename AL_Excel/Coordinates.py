""" AL_Excel.Coordinates

    Implementation of a Coordinate Class which makes it easier to handle locations
"""
## Super Module
from openpyxl import utils
## Builtin
import collections
import re


__all__ = ["Coordinate", "Index"]

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
    def __init__(self,row,column = False):
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
    def row(self):
        return self._row.value if self._row else None
    @property
    def column(self):
        return self._column.value if self._column else None
    def isabsolute(self):
        return self._column.absolute or self._row.absolute
    def toA1string(self, absolute = True):
        """ Returns the Coordinate in A1 format.

            By default, will output absolute dollar signs (ex.- A$1).
            If absolute is False, will output both row and column as relative.
        """
        if not absolute:
            return f"{utils.cell.get_column_letter(self._column.value)}{self._row.value}"
        return f'{"$" if self._row.absolute else ""}{utils.cell.get_column_letter(self._column.value)}{"$" if self._column.absolute else ""}{self._row.value}'
    def __add__(self,other):
        if isinstance(other,Coordinate):
            if (self._row.absolute and other._row.absolute)\
               or self._column.absolute and other._column.absolute:
                raise AttributeError(f"Cannot add two Absolute references: {self} + {other}")
            return Coordinate(row = addindices(self._row,other._row), column = addindices(self._column, other._column))
        ## Otherwise, try to convert to Coordinate
        ## Conversion will automatically handle Address Strings and 2-length iterables (i.e. (row,column) Tuple)
        try: return self + Coordinate(other)
        except: pass
    def __eq__(self,other):
        if isinstance(other,Coordinate):
            return self._row == other._row and self._column == other._column
    def __repr__(self):
        return f"{self.__class__}({self._row},{self._column})"


Index = collections.namedtuple("indextuple","type, value, absolute")

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

def converttotuple(value):
    if isinstance(value,str):
        regex = FULLRE.search(value)
        if regex:
            syntax = "RC" if regex.group("RC") else "A1"
            row = parsecoordregex(regex.group(f"{syntax}row"),'row')
            column = parsecoordregex(regex.group(f"{syntax}column"),'column')
            return row,column
        else:
            raise TypeError(f"Invalid coordinates: {value}")
    try: r,c = value
    except: raise TypeError(f"Cooridinates must be Coordinate-Formatted string or a tuple: {value}")
    ## Using list so that we can freely swap the indices when necessary
    indices = [parseindex(r),parseindex(c)]
    ###       I'm handling the two indices ambiguously rather than simply calling one "row"      ###
    ###  and the other "column" because assuming so causes complexity in the code to compensate  ###
    ################################################################################################
    ### Dear Past Self:                                                                          ###
    ###     Examples, please...                                                                  ###
    ### Sincerely, Future Self                                                                   ###

    ## Handle Non-Indexes
    nonindices = [index.value for index in indices if index.value is None]
    if nonindices:
        return handlenonindices(indices)

    ## row/column are ambiguous (because they were both a numeric index instead of Excel notation)
    if all(index.type is None for index in indices):
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

    ## Both Ambiguous is handled above
    ## Ambiguous Index is Column
    if row and not column:
        row = row[0]
        column = ambiguous[0]
        column = column._replace(type = "column")
    ## Ambiguous Index is Row
    elif column and not row:
        column = column[0]
        row = ambiguous[0]
        row = row._replace(type = "row")
    ## No Ambiguous indices
    else:
        row = row[0]
        column = column[0]
    return row,column

def handlenonindices(indices):
    """ Fork to handle missing coordinate values """
    nonindices = [index.value for index in indices if index.value is None]
    ## Both are None indexes (p.s.- >= is extraneous, but we're doing it anyway)
    if len(nonindices) >= 2:
        raise ValueError(f"No Indicies provided: {indices}")

    goodind = [i for i,index in enumerate(indices) if index.value is not None][0]
    goodindex = indices[goodind]
    ## If goodindex is ambiguous, then we need to set it's type
    if goodindex.type is None:
        ## type is determined by position in the tuple 
        ## i.e. (None,1) -> (NoneIndex,AmbiguousIndex==1) -> (NoneRowIndex,ColumnIndex==1)
        goodindex = goodindex._replace(type = COORDINATEORDER[goodind])
    ## Otherwise, we need to double-check goodindex's position in the tuple
    else:
        ## If goodindex is not in the correct position
        if goodindex.type != COORDINATEORDER[goodind]:
            ## Swap the coordinates
            indices[0],indices[1] = indices[1],indices[0]
    ## By this point the indices should read (row,None) or (None,column)
    return indices

def parseindex(value):
    """ Function for finding information about a specific descriptor """
    ## Handle Row/Column References ("$A" where one one index is None)
    if value is None:
        return Index(None,None,False)
    ## Check if coord is an int
    ## As it turns out, bool is a subclass of int :-/
    if isinstance(value,int) and not isinstance(value,bool):
        return Index(None,value,True)
    if isinstance(value,str):
        ## Check against Regex
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
        ## This module only uses True/False for absolute
        if value.absolute not in [True,False]: raise ValueError(f"Index's absolute value is not True or False: {value}")
        ## Index is valid
        return value
    ## Handle Tuples and Lists (possibly other iterables later)
    if isinstance(value,(tuple,list)):
        ## Only accept lengths 1,2
        if 0 > len(value) or len(value) > 2:
            raise ValueError(f"Coordinate part's length should be 1 or 2 (index,[absolute]): received length {len(value)}")
        index = parseindex(value[0])
        if len(value) == 2:
            absolute = value[1]
            ## Convert string
            if isinstance(absolute,str):
                absolute = absolute.lower()
                ## Acceptable absolute = True values
                if absolute in ["$","absolute"]: absolute = True
                ## Only absolute = False value
                elif absolute == "": absolute = False
                else: raise ValueError(f"Coordinate part's second index is an unknown string: {value[1]}")
            ## Acceptable values
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

def parsecoordregex(value,colrow):
    """ Ambidextrous helper function for parseindex """
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

def addindices(index1,index2):
    """ Adds two indices together. Both must have the same type and at least one must be relative (Index.absolute == False)"""
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
    if result.value <= 0:
        raise ValueError(f"Absolute Reference reduced below 0: {result}")
    return result