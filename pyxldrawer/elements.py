"""Drawing element classes"""

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from pandas import isnull
import sys, yaml
from collections import OrderedDict

###############################################################################

class Element(object):
    """Implementation of an atomic report element
    
    Attributes:
        value (any): cell value; may be of any atomic type
        height (int): height as a number of cells (rows); non-negative
        width (int): width as a number of cells (columns); non-negative
        style (xlsxwriter.format.Format): Element's style
        comment (str): comment text; defaults to none
        comment_params (dict): comment params (see xlsxwriters docs); defaults to {}
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def value(self):
        return self._value
    
    @value.setter
    def value(self, value):
        if isnull(value):
            value = ''
        self._value = value
    
    @property
    def height(self):
        """Element's height
        """
        return self._height
    
    @height.setter
    def height(self, value):
        if not isinstance(value, int):
            raise TypeError('height has to be a positive int')
        elif value <= 0:
            raise ValueError('height has to be > 0')
        self._height = value
    
    @property
    def width(self):
        """Element's width
        """
        return self._width
    
    @width.setter
    def width(self, value):
        if not isinstance(value, int):
            raise TypeError('width has to be a positive int')
        elif value <= 0:
            raise ValueError('height has to be a positive int')
        self._width = value
    
    @property
    def style(self):
        """Element's style
        """
        return self._style
    
    @style.setter
    def style(self, value):
        if isinstance(value, (dict, xlsxwriter.format.Format)):
            self._style = value
        else:
            raise TypeError('style has to be either a dict or an instance of xlsxwriter.format.Format.')
    
    @property
    def comment(self):
        """Comment text
        """
        return self._comment
    
    @comment.setter
    def comment(self, value):
        self._comment = value
    
    @property
    def comment_params(self):
        """Comment params dictionary
        """
        return self._comment_params
    
    @comment_params.setter
    def comment_params(self, value):
        if not isinstance(value, dict):
            raise TypeError('comment_params has to be a dict (it may be emtpy).')
        self._comment_params = value
    
    # -------------------------------------------------------------------------
    
    def __init__(self, value, height = 1, width = 1, style = {}, comment = None, comment_params = {}):
        """Constructor method
        """
        self.value = value
        self.height = height
        self.width = width
        self.style = style
        self.comment = comment
        self.comment_params = comment_params
    
    def make_style(self, wb):
        """Prepare Element's style for drawing
        
        Args:
            wb (xlsxwriter.workbook.Workbook): workbook to register a style in
        """
        if isinstance(self.style, dict):
            self.style = wb.add_format(self.style)
    
    def xl_upleft(self, x, y):
        """Get upper-left corner coordinates of the Element in the standard excel notation
        
        Args:
            x (int): x-coordinate of the upper-left corner of the Element
            y (int): y-coordinate of the upper-left corner of the Element
        
        Returns:
            str: string with an excel address of the upper-left corner
        """
        return xl_rowcol_to_cell(x, y)
    
    def xl_loright(self, x, y):
        """Get lower-right corner cooridnates of the Element in the standard excel notation
        """
        return xl_rowcol_to_cell(x + self.height - 1, y + self.width - 1)
    
    def xl_range(self, x, y):
        """Get range covered with the Element in the standard excel notation
        
        Returns:
            str: string representing an excel range covered by the Element
        """
        upleft = self.xl_upleft(x, y)
        loright = self.xl_loright(x, y)
        return upleft + ':' + loright
    
    def draw(self, x, y, ws, wb):
        """Draw Element in the worksheet
        
        Args:
            x (int): x-coordinate of the upper-left corner of the Element
            y (int): y-coordinate of the upper-left corner of the Element
            ws (xlsxwriter.worksheet.Worksheet): worksheet to write the Element in
            wb (xlsxwriter.workbook.Workbook): workbook the worksheet is in
        """
        self.make_style(wb)
        if self.width == 1 and self.height == 1:
            ws.write(x, y, self.value, self.style)
        else:
            rng = self.xl_range(x, y)
            ws.merge_range(rng, self.value, self.style)
        if self.comment is not None:
            addr = self.xl_upleft(x, y)
            ws.write_comment(addr, self.comment, self.comment_params)

###############################################################################

class HeaderElement(Element):
    """Header element class
    
    This is an extension of the Element class that adjust column widths while drawing.
    If col_width is None then column width is left untouched.
    If it is 'auto' then auto-resizing is done (adjusting to the length of the cell value text).
    If it is flaot then fixed width is set.
    
    Attributes:
        col_width (float/str/None): column width
        padding (float): padding addedd to both sides in auto-resizing
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def col_width(self):
        """Column width given as a float
        """
        return self._col_width
        
    @col_width.setter
    def col_width(self, value):
        if isinstance(value, str):
            if value != 'auto':
                raise ValueError("col_width has to be float, None or 'auto'.")
        elif value is None:
            pass
        else:
            try:
                value = float(value)
            except (TypeError, ValueError):
                raise TypeError("col_width has to be float, None or 'auto'.")
        self._col_width = value
        
    @property
    def padding(self):
        """Padding given as a float
        """
        return self._padding
    
    @padding.setter
    def padding(self, value):
        self._padding = float(value)
    
    # -------------------------------------------------------------------------
    
    def __init__(self, value, height = 1, width = 1, style = {}, 
                 comment = None, comment_params = {}, 
                 col_width = 'auto', padding = 1.0):
        """Constructor method
        """
        Element.__init__(self, value, height, width, style, comment, comment_params)
        self.col_width = col_width
        self.padding = padding
    
    def _value_len(self):
        """Computes length of the element's value
        """
        if self.value is not None:
            return len(str(self.value))
        else:
            return None
    
    def draw(self, x, y, ws, wb):
        """Extension of the draw method of the parent class
        """        
        Element.draw(self, x, y, ws, wb)
        if isinstance(self.col_width, float):
            col_width = self.col_width
        elif isinstance(self.col_width, str) and self.col_width == 'auto':
            try:
                col_width = float(self._value_len() + self.padding * 2) / self.width
            except TypeError:
                return
        elif self.col_width is None:
            return
        else:
            raise ValueError('incorrect value of col_width.')
        ws.set_column(y, y + self.width - 1, col_width)

###############################################################################

class Matrix(object):
    """Matrix of elements
    
    Useful for drawing rows, columns and matrices / tables.
    It provides easy means for defining borders of areas in an excel worksheet.
    
    Attributes:
        matrix (dict): matrix of elements
        nrow (int): number of rows
        ncol (int): number of columns
        height (int): height
        width (int): width
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def matrix(self):
        """Matrix of elements
        """
        return self._matrix
    
    @matrix.setter
    def matrix(self, value):
        if not isinstance(value, dict):
            raise TypeError('matrix has to be a dict.')
        self._matrix = value
        self._nrow = self._count_rows()
        self._ncol = self._count_cols()
        
    
    @property
    def nrow(self):
        """Number of rows
        """
        return self._nrow
    
    @nrow.setter
    def nrow(self, value):
        raise AttributeError('nrow can not be manually set.')
    
    @property
    def ncol(self):
        """Number of columns
        """
        return self._ncol
    
    @ncol.setter
    def ncol(self, value):
        raise AttributeError('ncol can not be manually set.')
        
    @property
    def height(self):
        return self._height
    
    @height.setter
    def height(self, value):
        if not isinstance(value, int):
            raise TypeError('height has to be an int.')
        if value < 1:
            raise ValueError('height has to be positive.')
        self._height = value
    
    @property
    def width(self):
        return self._width
    
    @width.setter
    def width(self, value):
        if not isinstance(value, int):
            raise TypeError('width has to be an int.')
        if value < 1:
            raise ValueError('width has to be positive.')
        self._width = value
    
    # -------------------------------------------------------------------------
        
    def __init__(self, values, height = 1, width = 1, style = {}, 
                            comment = None, comment_params = {},
                            col_width = None, padding = 1.0,
                            top = {}, right ={}, bottom = {}, left = {}):
        """Constructor method
        
        Args:
            values (list/dict): values matrix or list of lists of row values
            height (int): height of cells
            width (int): width of cells
            style (list/dict): style dict or style matrix (dict of style dicst) or list of lists
            comment (list/str/dict): comment text or matrix of comment texts or a list of lists
            comment_params (list/dict): comment params dict or matrix of comment params (dict of dicts) or a list of lists
            col_width (float): col_width to set; defaults to None which makes no adjustment
            padding (float): padding to ad if col_width = 'auto'
            top (dict): additional styling for top border
            right (dict): additional styling for right border
            bottom (dict): additional styling for bottom border
            left (dict): additional styling for right border
        """
        if isinstance(values, list):
            values = self.lists_to_matrix(values)
        if isinstance(height, list):
            height = self.lists_to_matrix(height)
        if isinstance(width, list):
            width = self.lists_to_matrix(width)
        if isinstance(style, list):
            style = self.lists_to_matrix(style)
        if isinstance(comment, list):
            comment = self.lists_to_matrix(comment)
        if isinstance(comment_params, list):
            comment_params = self.lists_to_matrix(comment_params)
        self.make_element_matrix(values, height, width, style, comment, comment_params, col_width, padding)
        self._nrow = self._count_rows()
        self._ncol = self._count_cols()
        self.height = self.nrow * height
        self.width = self.ncol * width
        
        # Add border styles ---
        for elem in self.border(which = 't'):
            elem.style = self._merge_styles(elem.style, top)
        for elem in self.border(which = 'r'):
            elem.style = self._merge_styles(elem.style, right)
        for elem in self.border(which = 'b'):
            elem.style = self._merge_styles(elem.style, bottom)
        for elem in self.border(which = 'l'):
            elem.style = self._merge_styles(elem.style, left)
    
    def _merge_styles(self, style, additional_style):
        """Add and/or change styling dict
        
        Args:
            style (dict): original style dictionary
            additional_style (dict): dict with additional styling rules
        
        Returns:
            dict: merge styling dictionary
        """
        merged_style = style.copy()
        for key, value in additional_style.items():
            merged_style[key] = value
        return merged_style
        
    def _count_rows(self, matrix = None):
        if matrix is None:
            matrix = self.matrix
        return max([ x[0] for x in matrix.keys() ]) + 1
    
    def _count_cols(self, matrix = None):
        if matrix is None:
            matrix = self.matrix
        return max([ x[1] for x in matrix.keys() ]) + 1
    
    def get(self, x, y):
        """Get matrix element by index
        
        Args:
            x (int): row index
            y (int): column index
        
        Returns:
            Element
        """
        if not isinstance(x, int) or not isinstance(y, int):
            raise TypeError('indices must be integers.')
        if x < 0 or x > self.nrow - 1:
            raise IndexError('x index out of range.')
        if y < 0 or y > self.ncol - 1:
            raise IndexError('y index out of range.')
        return self.matrix[(x, y)]
    
    def set(self, x, y, value):
        """Set matrix element by index
        
        Args:
            value (Element): object inheriting from the Element class
        """
        if not issubclass(value, Element):
            raise TypeError('value has to inherit from the Element class.')
        if not isinstance(x, int) or not isinstance(y, int):
            raise TypeError('indices must be integers.')
        if x < 0 or x > self.nrow - 1:
            raise IndexError('x index out of range.')
        if y < 0 or y > self.ncol - 1:
            raise IndexError('y index out of range.')
        self.matrix[(x, y)] = value
    
    def border(self, which, corner1 = True, corner2 = True):
        """Get border of the matrix
        
        Args:
            which (str): t, r, b or l (top, right, bottom, left)
            corner1 (bool): whether to include top/left corner element
            corner2 (bool): whether to include bottom/right element
        
        Returns:
            list: list of border elements in the proper order
        """
        if which == 't' or which == 'b':
            n = 0 if which == 't' else self.nrow - 1
            m0 = 0 if corner1 else 1
            m1 = self.ncol if corner2 else self.ncol - 1
            return [ self.get(n, x) for x in range(m0, m1) ]
        elif which == 'r' or which == 'l':
            n0 = 0 if corner1 else 1
            n1 = self.nrow if corner2 else self.nrow - 1
            m = 0 if which == 'l' else self.ncol - 1
            return [ self.get(x, m) for x in range(n0, n1)]
        else:
            raise ValueError("which has to be either 't', 'r', 'b' or 'l'.")
    
    def corner(self, which):
        """Get corner element of a matrix
        
        Args:
            which (int/str): 1/topleft, 2/topright, 3/bottomright, 4/bottomleft
        
        Returns:
            Element: corner element
        """
        vals = {1: 'topright', 2: 'bottomright', 
                3: 'bottomleft', 4: 'topleft', }
        if isinstance(which, str):
            pass
        elif isinstance(which, int):
            which = vals[which]
        else:
            raise TypeError('which has to be a str or int.')
        if which == 'topleft':
            return self.matrix[(0, 0)]
        elif which == 'topright':
            return self.matrix[(0, self.ncol - 1)]
        elif which == 'bottomright':
            return self.matrix[(self.nrow - 1, self.ncol - 1)]
        elif which == 'bottomleft':
            return self.matrix[(self.nrow - 1, 0)]
        else:
            raise ValueError("which has to be either 1/'topright', 2/'bottomright', 3/'bottomleft' or 4/'topleft'")
        
    def lists_to_matrix(self, L):
        """Converts a list of lists to a matrix
        
        Args:
            L: list of lists with row values / params / objects
        
        Returns:
            dict: matrix in a form of a dictionary
        """
        matrix = {}
        lengths = [ len(x) for x in L ]
        max_m = max(lengths)
        min_m = min(lengths)
        if max_m == min_m:
            m = max_m
        else:
            raise ValueError('Sub lists do not have identical lenghts.')
        n = len(L)
        for i in range(n):
            for j in range(m):
                matrix[(i, j)] = L[i][j]
        return matrix
    
    def make_element_matrix(self, values, height = 1, width = 1, style = {}, 
                            comment = None, comment_params = {},
                            col_width = None, padding = 3.0):
        """Make element matrix from matrices of values, height etc.
        
        Args:
            values (dict): values matrix
            height (int/dict): height of cells or height matrix
            width (int/dict): width of cells or width matrix
            style (dict): style dict or style matrix (dict of style dicst)
            comment (str/dict): comment text or matrix of comment texts
            comment_params (dict): comment params dict or matrix of comment params (dict of dicts)
            col_width (float): col_width to set; defaults to None which makes no adjustment
            padding (float): padding to ad if col_width = 'auto'
        """
        matrix = {}
        n = self._count_rows(values)
        m = self._count_cols(values)
        for i in range(n):
            for j in range(m):
                elem = HeaderElement(
                    value = values[(i, j)],
                    height = height[(i, j)] if isinstance(height, dict) else height,
                    width = width[(i, j)] if isinstance(width, dict) else width,
                    style = style[(i, j)] if len(style) > 0 and all([ isinstance(x, dict) for x in style.values() ]) else style,
                    comment = comment[(i, j)] if comment is not None and len(comment) > 0 and isinstance(comment, dict) else comment,
                    comment_params = comment_params[(i, j)] if len(comment_params) > 0 and all([ isinstance(x, dict) for x in comment_params.values() ]) else comment_params,
                    col_width = col_width,
                    padding = padding
                )
                matrix[(i, j)] = elem
        self.matrix = matrix
        
    def draw(self, x, y, ws, wb):
        """Draw Matrix object in a worksheet
        
        Args:
            x (int): x-coordinate (rows)
            y (int): y-coordinate (columns)
            ws (xlsxwriter.worksheet.Worksheet): worksheet to draw in
            wb (xlsxwriter.workbook.Workbook): workbook the worksheet is in
        """
        y0 = y
        for i in range(self.nrow):
            height = 1
            for j in range(self.ncol):
                elem = self.get(i, j)
                elem.draw(x, y, ws, wb)
                y += elem.width
                if elem.height > height:
                    height = elem.height
            y = y0
            x += height
            
###############################################################################

class TreeElement(object):
    """Element with a row of sub elements
    
    Attributes:
        height (int): total height of the element
        width (int): total width of the element
        parent (Element): object inheriting from the Element class
        children (LineElement): object inheriting from the LineElement class
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def width(self):
        return self._width
    
    @width.setter
    def width(self, value):
        if not isinstance(value, int):
            raise TypeError('width has to be an int.')
        elif value < 1:
            raise ValueError('width has to be positive.')
        self._width = value
    
    @property
    def height(self):
        return self._height
    
    @height.setter
    def height(self, value):
        if not isinstance(value, int):
            raise TypeError('height has to be an int.')
        elif value < 1:
            raise ValueError('height has to be positive.')
        self._height = value
    
    @property
    def parent(self):
        return self._parent
    
    @parent.setter
    def parent(self, value):
        if not issubclass(type(value), Element):
            raise TypeError('parent has to inherit from Element class.')
        self._parent = value
    
    @property
    def children(self):
        return self._children
    
    @children.setter
    def children(self, value):
#        if not issubclass(type(value), Matrix):
#            raise TypeError('children has to inherit from Matrix class.')
#        if value.nrow != 1:
#            raise ValueError('children has to have only one row.')
        self._children = value
    
    # -------------------------------------------------------------------------
    
    def __init__(self, parent, children):
        """Constructor method
        
        Args:
            parent (Element): object inheriting from Element class
            children (Matrix): object inheriting from Matrix class; has to have only 1 row
        """
        if parent.width != children.width:
            raise ValueError('parent and children widths are not the same.')
        self.parent = parent
        self.children = children
        self.height = parent.height + children.height
        self.width = parent.width
    
    def draw(self, x, y, ws, wb):
        """Drawing method
        
        Args:
            x (int): x-coordinate
            y (int): y-coordinate
            ws (xlsxwriter.worksheet.Worksheet): worksheet to draw in
            wb (xlsxwriter.workbook.Workbook): workbook the worksheet is in
        """
        self.parent.draw(x, y, ws, wb)
        self.children.draw(x + self.parent.height, y, ws, wb)

###############################################################################

class LineElement(object):
    """Horizontal or vertical line of elements
    
    Attributes:
        height (int): total height of element / height of the heighes element
        width (int): width of the widest element / total width of the elements
        vertical (bool): is the line vertical or horizontal
        elements (list): list of objects equiped with the 'draw' method
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def height(self):
        return self._height
    
    @height.setter
    def height(self, value):
        if not isinstance(value, int) or value < 1:
            raise TypeError('height has to be a positive integer.')
        self._height = value
    
    @property
    def width(self):
        return self._width
    
    @width.setter
    def width(self, value):
        if not isinstance(value, int) or value < 1:
            raise TypeError('width has to be a positive integer.')
        self._width = value
    
    @property
    def vertical(self):
        return self._vertical
    
    @vertical.setter
    def vertical(self, value):
        if not isinstance(value, bool):
            raise TypeError('vertical has to be a bool.')
        self._vertical = value
    
    @property
    def elements(self):
        return self._elements
    
    @elements.setter
    def elements(self, value):
        if not isinstance(value, list):
            raise TypeError('elements has to be a list.')
        self._elements = value
    
    # -------------------------------------------------------------------------
    
    def __init__(self, elements, vertical = False):
        """Constructor method
        """
        self.elements = elements
        self.vertical = vertical
        if vertical:
            self.height = sum([ x.height for x in elements ])
            self.width = max([ x.width for x in elements ])
        else:
            self.height = max([ x.height for x in elements ])
            self.width = sum([ x.width for x in elements ])
    
    def draw(self, x, y, ws, wb):
        """Draw LineElements in a worksheet
        
        Args:
            x (int): x-coordinate
            y (int): y-coordinate
            ws (xlsxwriter.worksheet.Worksheet): worksheet to draw in
            wb (xlsxwriter.workbook.Workbook): workbook the worksheet is in
        """
        if self.vertical:
            for elem in self.elements:
                elem.draw(x, y, ws, wb)
                x += elem.height
        else:
            for elem in self.elements:
                elem.draw(x, y, ws, wb)
                y += elem.width

###############################################################################

class Dictionary(object):
    """Visual/tabular representaion of a key => value set
    
    This class implements a layout of fields in a report,
    in which there is one column (a key column)
    separated from a second column by a horizontal space of a given width
    that presents key (titles) and the second column presents content (values)
    for given keys. Useful form making into/definitions pages for various reports.
    
    Attributes:
        structure (OrderedDict/str): definition of the structure of a Dictionary (key => value) or a path to the .yaml config file
        hspace (int): width of the horizontal space between key column and value column
        vspace (int): default vertical spacing between fields
        text_params (dict): key and values to be interpolated in text
        field_params (dict): default set of params passed to the HeaderElement constructor (field column) as **kwargs
        content_params (dict): default set of params passed to the HeaderElement constructor (content column) as **kwargs
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def structure(self):
        return self._structure
    
    @structure.setter
    def structure(self, value):
        if not isinstance(value, (OrderedDict, str)):
            raise TypeError('structure has to be an OrderedDict.')
        if isinstance(value, str):
            self._structure = self.load_config(value)
        else:
            self._structure = value
    
    @property
    def hspace(self):
        return self._hspace
    
    @hspace.setter
    def hspace(self, value):
        if not isinstance(value, int):
            raise TypeError('hspace has to be an int.')
        self._hspace = value
    
    @property
    def vspace(self):
        return self._vspace
    
    @vspace.setter
    def vspace(self, value):
        if not isinstance(value, int):
            raise TypeError('vspace has to be an int.')
        self._vspace = value
    
    @property
    def text_params(self):
        return self._text_params
    
    @text_params.setter
    def text_params(self, value):
        if not isinstance(value, dict):
            raise TypeError('text_params has to be a dict.')
        self._text_params = value
    
    @property
    def field_params(self):
        return self._field_params
    
    @field_params.setter
    def field_params(self, value):
        if not isinstance(value, dict):
            raise TypeError('field_params has to be a dict.')
        self._field_params = value
    
    @property
    def content_params(self):
        return self._content_params
    
    @content_params.setter
    def content_params(self, value):
        if not isinstance(value, dict):
            raise TypeError('content_params has to be a dict.')
        self._content_params = value
    
    @property
    def height(self):
        return self._height
    
    @height.setter
    def height(self, value):
        if not isinstance(value, int):
            raise TypeError('height has to be a positive int.')
        elif value < 1:
            raise ValueError('height has to be a positive int.')
        self._height = value
    
    @property
    def width(self):
        return self._width
    
    @width.setter
    def width(self, value):
        if not isinstance(value, int):
            raise TypeError('width has to be a positive int.')
        elif value < 1:
            raise ValueError('width has to be a positive int.')
        self._width = value
    
    # -------------------------------------------------------------------------
    
    def __init__(self, structure, hspace = 1, vspace = 0, text_params = {},
                 field_params = {}, content_params = {}):
        """Constructor method
        """
        self.structure = structure
        self.hspace = hspace
        self.vspace = vspace
        self.text_params = text_params
        self.field_params = field_params
        if content_params.get('col_width') is None:
            content_params['col_width'] = None
        self.content_params = content_params
        
        # Determine height and width ---
        height = 0
        width = 0
        for field, content in self.structure.items():
            fh = field_params.get('height', 1)
            fw = field_params.get('width', 1)
            cw = content_params.get('width', 1)
            w = fw + cw + self.vspace
            if w > width:
                width = w
            vals = content['content']
            if not isinstance(vals, list):
                vals = [vals]
            ch = len(vals) * content_params.get('height', 1)
            if fh > ch:
                height += fh
            else:
                height += ch
        self.height = height
        self.width = width
    
    def interpolate_string(self, s):
        """Interpolate a string using text_params
        
        Args:
            s (str): a string
        """
        return s.format(f = self.text_params)
    
    def load_config(self, path = None):
        """Loads config from a config.yaml file
    
        Args:
            path (str): path to a config file; may be None, then Collector object's default is used
            
        Returns:
            OrderedDict: config parsed to a dictionary
        """        
        if path is None:
            path = self.config_path
    
        def ordered_load(stream, Loader = yaml.Loader, object_pairs_hook = OrderedDict):
            class OrderedLoader(Loader):
                pass
            def construct_mapping(loader, node):
                loader.flatten_mapping(node)
                return object_pairs_hook(loader.construct_pairs(node))
            OrderedLoader.add_constructor(
                yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG,
                construct_mapping
            )
            return yaml.load(stream, OrderedLoader)    
        cnf = open(path, 'r')
        try:
            config = ordered_load(cnf)
        except yaml.YAMLError as exc:
            sys.exit(exc)
        finally:
            cnf.close()
        return config   

    def _merge_styles(self, style, additional_style):
        """Add and/or change styling dict
        
        Args:
            style (dict): original style dictionary
            additional_style (dict): dict with additional styling rules
        
        Returns:
            dict: merge styling dictionary
        """
        merged_style = style.copy()
        for key, value in additional_style.items():
            merged_style[key] = value
        return merged_style

    def draw(self, x, y, ws, wb):
        """Draw Dictionary in a worksheet
        
        Args:
            x (int): x-coordinate (rows)
            y (int): y-coordinate (columns)
            ws (xlsxwriter.worksheet.Worksheet): worksheet to draw in
            wb (xlsxwriter.workbook.Workbook): workbook to draw in
        """
        y0 = y
        for field, data in self.structure.items():
            field_params = self._merge_styles(self.field_params, data.get('field_params', {}))
            content_params = self._merge_styles(self.content_params, data.get('content_params', {}))
            field_value = self.interpolate_string(field)
            vspace = data.get('vspace', self.vspace)
            Field = HeaderElement(field_value, **field_params)
            Field.draw(x, y, ws, wb)
            content = data['content']
            if not isinstance(content, list):
                content = [content]
            elif content is None:
                content = ['']
            for value in content:
                if isinstance(value, str):
                    value = self.interpolate_string(value)
                Content = HeaderElement(value, **content_params)
                Content.draw(x, y  + Field.width + self.hspace, ws, wb)
                x += Content.height
            y = y0
            x += vspace

###############################################################################