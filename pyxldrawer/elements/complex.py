"""Complex drawing element classes"""

from basic import Element, HeaderElement
from matrix import Matrix


class TreeElement(object):
    """Element with a row of sub elements
    
    Attributes:
        height (int): total height of the element
        width (int): total width of the element
        parent (Element): object inheriting from the Element class
        children (LineElement): object inheriting from the LineElement class
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
#        if not issubclass(type(value), Element):
#            raise TypeError('parent has to inherit from Element class.')
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


class LineElement(object):
    """Horizontal or vertical line of elements
    
    Attributes:
        height (int): total height of element / height of the heighes element
        width (int): width of the widest element / total width of the elements
        vertical (bool): is the line vertical or horizontal
        elements (list): list of objects equiped with the 'draw' method
    """
    
    ###########################################################################
    
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
    
    ###########################################################################
    
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
                y += elem.height
        else:
            for elem in self.elements:
                elem.draw(x, y, ws, wb)
                x += elem.width