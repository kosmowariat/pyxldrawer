"""Complex drawing element classes"""

from basic import Element, HeaderElement
from matrix import Matrix


class TreeElement(object):
    """Element with a row of sub elements
    
    Attributes:
        height (int): total height of the element
        width (int): total width of the element
        parent (Element): object inheriting from the Element class
        children (Matrix): object inheriting from the Matrix class; must have nrow == 1.
    """
    
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
        if not issubclass(value, Element):
            raise TypeError('parent has to inherit from Element class.')
        self._parent = value
    
    @property
    def children(self):
        return self._children
    
    @children.setter
    def children(self, value):
        if not issubclass(value, Matrix):
            raise TypeError('children has to inherit from Matrix class.')
        if value.nrow != 1:
            raise ValueError('children has to have only one row.')
        self._children = value
    
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
        self.children.draw(x, y + self.parent.height, ws, wb)