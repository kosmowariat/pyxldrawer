"""Standard drawing elements of pyxldrawer"""

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

#def val2text(val):
#            if val is None:
#                return ''
#            elif isinstance(val, str):
#                return val
#            else:
#                return str(val)

class Element(object):
    """Implementation of an atomic report element
    
    Attributes:
        value (any): cell value; may be of any atomic type
        _height (int): height as a number of cells (rows); non-negative
        _width (int): width as a number of cells (columns); non-negative
        _style (xlsxwriter.format.Format): Element's style
        comment (str): comment text; defaults to none
        comment_params (dict): comment params (see xlsxwriters docs); defaults to {}
    """
    
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
            wb.add_format(self.style)
    
    def xl_upleft(self, x, y):
        """Get upper-left corner coordinates of the Element in the standard excel notation
        
        Args:
            x (int): x-coordinate of the upper-left corner of the Element
            y (int): y-coordinate of the upper-left corner of the Element
        
        Returns:
            str: string with an excel address of the upper-left corner
        """
        return xl_rowcol_to_cell(y, x)
    
    def xl_loright(self, x, y):
        """Get lower-right corner cooridnates of the Element in the standard excel notation
        """
        return xl_rowcol_to_cell(y + self.height - 1, x + self.width - 1)
    
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
        if self.x == 1 and self.y == 1:
            ws.write(x, y, self.value, self.style)
        else:
            rng = self.xl_range(x, y)
            ws.merge_range(rng, self.value, self.style)
        if self.comment is not None:
            addr = self.xl_upleft(x, y)
            ws.write_comment(addr, self.comment, self.comment_params)

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
    def __init__(self, value, height = 1, width = 1, style = 'Normal', col_width = 'auto', padding = 3.0, comment = None):
        """Constructor method
        """
        Element.__init__(self, value, height, width, style, comment)
        self.col_width = float(col_width) if isinstance(col_width, int) else col_width
        self.padding = float(padding) if isinstance(col_width, int) else padding
    
    def draw(self, x, y, ws):
        """Extension of the draw method of the parent class
        """
        # Import modules ---
        from openpyxl.utils import get_column_letter as column
        
        Element.draw(self, x, y, ws)
        if isinstance(self.col_width, float):
            col_width = self.col_width
        elif isinstance(self.col_width, str) and self.col_width == 'auto':
            text_val = val2text(self.value)
            col_width = float(len(text_val) + self.padding * 2) / self.width
        if isinstance(col_width, float):
            for i in range(x + 1, x + 1 + self.width):
                ws.column_dimensions[column(i)].width = col_width
            
class ElementRow(object):
    """Row of multiple elements
    
    Attributes:
        elements (list): list of object inheriting from the Element class
        height (int): height of the row
        width (int): width of the row
    """
    def __init__(self, elements):
        """Constructor method
        """
        self.height = max([ elem.height for elem in elements ])
        self.width = sum([ elem.width for elem in elements ])
        self.elements = elements
    
    def draw(self, x, y, ws):
        """Draw ElementRow in a worksheet
        
        Args:
            x (int): x-coordinate of the upper-left corner of the ElementRow
            y (int): y-coordinate of the upper-left corner of the ElementRow
        """
        for elem in self.elements:
            elem.draw(x, y, ws)
            x += elem.width

class ElementColumn(object):
    """Column of multiple elemnts
    """
    def __init__(self, elements):
        """Constructor method
        """
        self.height = sum([ elem.height for elem in elements ])
        self.width = max([ elem.width for elem in elements ])
        self.elements = elements
    
    def draw(self, x, y, ws, adjust_width = False, padding = 0.0):
        """Draw ElementRow in a worksheet
        
        Args:
            x (int): x-coordinate of the upper-left corner of the ElementRow
            y (int): y-coordinate of the upper-left corner of the ElementRow
        """
        for elem in self.elements:
            elem.draw(x, y, ws)
            y += elem.height
        if adjust_width:
             from openpyxl.utils import get_column_letter as column
             maxlen = float(max([ len(val2text(elem.value)) for elem in self.elements ]))
             maxwidth = max([ elem.width for eleme in self.elements ])
             col_width = maxlen + padding * 2
             for i in range(x + 1, x + 1 + maxwidth):
                 ws.column_dimensions[column(i)].width = col_width

def build_row(values, height, widths, styles = 'Normal'):
    """Build simple row of ordinary Element objects
    
    Args:
        values (list): list of values
        height (int): height of the row
        widths (int or list): width of the elements or a list of widths of the elements
        styles (str or NamedStyle or list): single name of a style or a NamedStyle object or a list of style names or NamedStyle objects
    
    Returns:
        ElementRow: simple ElementRow object
    """
    # Import modules ---
    from openpyxl.styles import NamedStyle
    
    n = len(values)
    if isinstance(widths, int):
        widths = [ widths for x in range(n) ]
    if isinstance(styles, str) or isinstance(styles, NamedStyle):
        styles = [ styles for x in range(n) ]
    row = []
    for data in zip(values, widths, styles):
        value = data[0]
        width = data[1]
        style = data[2]
        elem = Element(value, height, width, style)
        row.append(elem)
    return ElementRow(row)

def build_header_row(values, height, widths, styles = 'Normal', col_widths = 'auto', padding = 3.0):
    """Build header row made of HeaderElements object
    
    Works exactly the same as build_simple_row, but uses two extra arguments: col_widths and padding.
    """
    # Import modules ---
    from openpyxl.styles import NamedStyle
    
    n = len(values)
    if isinstance(widths, int):
        widths = [ widths for x in range(n) ]
    if isinstance(styles, str) or isinstance(styles, NamedStyle):
        styles = [ styles for x in range(n) ]
    if isinstance(col_widths, str) or isinstance(col_widths, float) or isinstance(col_widths, int):
        col_widths = [ col_widths for x in range(n) ]
    row = []
    for data in zip(values, widths, styles, col_widths):
        value = data[0]
        width = data[1]
        style = data[2]
        col_width = data[3]
        elem = HeaderElement(value, height, width, style, col_width, padding)
        row.append(elem)
    return ElementRow(row)

def build_column(values, heights, width, styles = 'Normal'):
    """Build simple column of ordinary Element objects
    
    Args:
        values (list): list of values
        heights (int or list): height of the elemebts or a list of heights of the elements
        width (int or list): width of the elements
        styles (str or NamedStyle or list): single name of a style or a NamedStyle object or a list of style names or NamedStyle objects
    
    Returns:
        ElementColumn: simple ElementColumn object
    """
    # Import modules ---
    from openpyxl.styles import NamedStyle
    
    n = len(values)
    if isinstance(heights, int):
        heights = [ heights for x in range(n) ]
    if isinstance(styles, str) or isinstance(styles, NamedStyle):
        styles = [ styles for x in range(n) ]
    col = []
    for data in zip(values, heights, styles):
        value = data[0]
        height = data[1]
        style = data[2]
        elem = Element(value, height, width, style)
        col.append(elem)
    return ElementColumn(col)

class TreeElement(object):
    """Element with child Elements below it
    
    Attributes:
        parent (Element): parent element (plain Element object)
        children (ElementRow): row of children elements represented by an ElementRow object
        height (int): height of the entire TreeElement
        width (int): width of the entire TreeElement
    """
    def __init__(self, parent, children):
        """Constructor method
        """
        self.parent = parent
        self.children = children
        self.height = parent.height + children.height
        self.width = max([parent.width, children.width])
    
    def draw(self, x, y, ws):
        """Draw TreeElement in a worksheet
        """
        self.parent.draw(x, y, ws)
        self.children.draw(x, y + self.parent.height, ws)

