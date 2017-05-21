"""Base drawing element classes"""

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

class Element(object):
    """Implementation of an atomic report element
    
    Attributes:
        value (any): cell value; may be of any atomic type
        _height (int): height as a number of cells (rows); non-negative
        _width (int): width as a number of cells (columns); non-negative
        _style (xlsxwriter.format.Format): Element's style
        _comment (str): comment text; defaults to none
        _comment_params (dict): comment params (see xlsxwriters docs); defaults to {}
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
        if x == 1 and y == 1:
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
        _col_width (float/str/None): column width
        _padding (float): padding addedd to both sides in auto-resizing
    """
    
    @property
    def col_width(self):
        """Column width given as a float
        """
        return self._col_width
        
    @col_width.setter
    def col_width(self, value):
        self._col_width = float(value)
        
    @property
    def padding(self):
        """Padding given as a float
        """
        return self._padding
    
    @padding.setter
    def padding(self, value):
        self._padding = float(value)
    
    def __init__(self, value, height = 1, width = 1, style = {}, 
                 comment = None, comment_params = {}, 
                 col_width = 'auto', padding = 3.0):
        """Constructor method
        """
        Element.__init__(self, value, height, width, style, comment, comment_params)
        self.col_width = col_width
        self.padding = padding
    
    def value_len(self):
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
                col_width = float(self.value_len() + self.padding * 2) / self.width
            except TypeError:
                return
        elif self.col_width is None:
            return
        else:
            raise ValueError('incorrect value of col_width.')
        ws.set_column(x, x + self.width - 1, col_width)
