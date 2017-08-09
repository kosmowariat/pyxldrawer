"""The main drawing controller class"""

import xlsxwriter
import re
from collections import OrderedDict
from xlsxwriter.utility import xl_rowcol_to_cell

###############################################################################
    
class Drawer(object):
    """Elements drawer
    
    This is an implementation of the drawing class.
    Drawer object is used for drawing actual drawing element in a .xlsx report.
    Its mechanics are quite simple: Drawer has its position in a standard cartesiax xy coordinate system
    and it can be fed with drawing elements which it in turn draws (according to their attributes)
    in a place it is currently located in.
    
    Drawers position are defined in matrix-like terms.
    X is rows (vertical dimension) and Y is column (horizontal dimension).
    
    Attributes:
        x (int): current x-coordinate (rows)
        y (int): current y-coordinate (columns)
        ws (xlsxwriter.worksheet.Workshet): worksheet to draw on
        wb (xlsxwriter.workbook.Workbook): workbook the worksheet in in
        height (int): height of the last drawed object
        width (int): width of the last drawed object
        prev_x (list): list of previous x-coordinates
        prev_y (list): list of previous y-coordinates
        checkpoints (OrderedDict): set of checkpoints
    """
    
    # -------------------------------------------------------------------------
    
    @property
    def x(self):
        """x coordinate
        """
        return self._x
    @x.setter
    def x(self, value):
        if not isinstance(value, int):
            raise TypeError('x coordinate has to be a non-negative int.')
        elif value < 0:
            raise ValueError('x coordinate has to be >= 0.')
        self._x = value
    
    @property
    def y(self):
        """y coordinate
        """
        return self._y
    @y.setter
    def y(self, value):
        if not isinstance(value, int):
            raise TypeError('y coordinate has to be a non-negative int.')
        elif value < 0:
            raise ValueError('y coordinate has to be >= 0.')
        self._y = value
    
    @property
    def ws(self):
        """Drawer's worksheet
        """
        return self._ws
    @ws.setter
    def ws(self, value):
        if not isinstance(value, xlsxwriter.worksheet.Worksheet):
            raise TypeError('ws has to be an instance of xlsxwriter.worksheet.Worksheet.')
        self._ws = value
    
    @property
    def wb(self):
        """Drawer's workbook
        """
        return self._wb
    @wb.setter
    def wb(self, value):
        if not isinstance(value, xlsxwriter.workbook.Workbook):
            raise TypeError('wb has to be an instance of xlsxwriter.workbook.Workbook.')
        self._wb = value
    
    # -------------------------------------------------------------------------
    
    def __init__(self, ws, wb, x = 0, y = 0):
        """Constructor method
        """
        self.x = x
        self.y = y
        self.ws = ws
        self.wb = wb
        self.height = 0
        self.width = 0
        self.checkpoints = OrderedDict()
        self.prev_x = []
        self.prev_y = []
    
    def __str__(self):
        """String representation of a Drawer object
        """
        text = object.__repr__(self)
        text += '\n\tx: ' + str(self.x)
        text += '\n\ty: ' + str(self.y)
        text += '\n\theight: ' + str(self.height)
        text += '\n\twidth: ' + str(self.width)
        text += '\n\tcheckpoints: ' + str(len(self.checkpoints))
        text += '\n\tprevious positions: ' + str(len(self.prev_x))
        text += '\n\tworksheet: ' + str(self.ws)
        return text
    
    def draw(self, elem, **kwargs):
        """Draw an element in a worksheet
        
        Args:
            elem (any): any object with a proper .draw() method
            **kwargs: keyword arguments passed to the invoked draw method
        """
        elem.draw(self.x, self.y, self.ws, self.wb, **kwargs)
        self.height = elem.height
        self.width = elem.width
    
    def move(self, x = 0, y = 0, back = False):
        """Move drawer
        
        Move drawer by specifed number of cells horizontally and/or vertically
        
        Args:
            x (int): number of cells to move in the x-direction (horizontal)
            y (int): number of cells to move in the y-direction (vertical)
            back (bool): whether to move forward or backward
        """
        self.prev_x.append(self.x)
        self.prev_y.append(self.y)
        if back:
            self.x -= x
            self.y -= y
        else:
            self.x += x
            self.y += y
    
    def move_horizontal(self, y = None, back = False):
        """Move drawer horizontally
        
        This method is useful, since it defaults to the last's object width.
        
        Args:
            y (int): number of cells to move; defaults to the width of the last drawed object
            back (bool): whether to move forward or backward
        """
        if y is None:
            y = self.width
        self.move(0, y, back = back)
    
    def move_vertical(self, x = None, back = False):
        """Move drawer vertically
        
        Defaults to the last object's height.
        
        Args:
            x (int): number of cells to move; default to the height of the last drawed object
            back (bool): whether to move forward or backward
        """
        if x is None:
            x = self.height
        self.move(x, 0, back = back)
    
    def add_checkpoint(self, name):
        """Adds current position as a checkpoint
        """
        self.checkpoints[name] = (self.x, self.y)
    
    def reset(self, checkpoint = None, x = 0, y = 0):
        """Reset Drawer position
        
        If checkpoint name (or index) is provided, then the Drawer is reset to the checkpoint.
        Otherwise it is reset to the given x and y coordinates.
        change_x and change_y flags are useful when a Drawer is to be reset to a checkpoint but only in one dimension.
        
        Args:
            name (str/int/None): name of a checkpoint to fall back to. Defaults to the origin (0, 0). Intgers are used as key indices.
            change_x/y (bool): whether x/y cooridnate should be changed
            x/y (int): new cooridnates to assig if checkpoint is None
        """
        self.prev_x.append(self.x)
        self.prev_y.append(self.y)
        
        if isinstance(checkpoint, str):
            if x is not None:
                self.x = self.checkpoints[checkpoint][0]
            if y is not None:
                self.y = self.checkpoints[checkpoint][1]
        elif isinstance(checkpoint, int):
            cp = self.checkpoints[list(self.checkpoints.keys()[checkpoint])]
            if x is not None:
                self.x = cp[0]
            if y is not None:
                self.y = cp[1]
        else:
            if x is not None:
                self.x = x
            if y is not None:
                self.y = y
        
    def fallback(self, n):
        """Fall back to nth previous step
        
        Args:
            n (int): number of steps to fall back. Negative values iterate from the historically first position.
        """
        self.reset(checkpoint = None, x = self.prev_x[-n], y = self.prev_y[-n], change_x = True, change_y = True)
    
    def xl_position(self, x = 0, y = 0):
        """Get Drawer's position
        
        Args:
            x (int): number of rows to shift when determining position
            y (int): number of columns to shift when determinin position
        
        Returns:
            str: string with an excel address of the upper-left corner
        """
        return xl_rowcol_to_cell(self.x + x, self.y + y)
    
    def xl_column(self, y = 0):
        """Get Drawer's current column in the excel notation
        
        Args:
            y (int): number of columns to shift when determining position
        """
        return re.sub('[0-9]', '', self.xl_position(y = y))
    
    def xl_row(self, x = 0):
        """Get Drawer's current row in the excel notation
        
        Args:
            x (int): number of rows to shift when determining position
        """
        return re.sub('[^0-9]', '', self.xl_position(x = x))
    
###############################################################################