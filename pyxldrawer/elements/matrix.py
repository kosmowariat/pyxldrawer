"""Matrix of Elements"""

from basic import Element, HeaderElement


class Matrix(object):
    """Matrix of elements
    
    Attributes:
        _matrix (dict): matrix of elements
        _nrow (int): number of rows
        _ncol (int): number of columns
        _height (int): height
        _width (int): width
    """
    
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
        
    def __init__(self, values, height = 1, width = 1, style = {}, 
                            comment = None, comment_params = {},
                            col_width = None, padding = 3.0,
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
        vals = {1: 'topleft', 2: 'topright', 3: 'bottomright', 4: 'bottomleft'}
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
            raise ValueError("which has to be either 1/'topleft', 2/'topright', 3/'bottomright' or 4/'bottomleft'")
        
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
            x (int): x-coordinate
            y (int): y-coordinate
            ws (xlsxwriter.worksheet.Worksheet): worksheet to draw in
            wb (xlsxwriter.workbook.Workbook): workbook the worksheet is in
        """
        x0 = x
        for i in range(self.nrow):
            height = 1
            for j in range(self.ncol):
                elem = self.get(i, j)
                elem.draw(x, y, ws, wb)
                x += elem.width
                if elem.height > height:
                    height = elem.height
            x = x0
            y += elem.height
            
