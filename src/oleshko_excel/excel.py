import openpyxl

class Excel:
    # items = []
    book = openpyxl.Workbook()
    sheet = book.active
    headings = False

    def __init__(self, file='Table', headings=False):
        self.file = file
        Excel.headings = headings

    def save(self):
        Excel.book.save(f'{self.file}.xlsx')
        Excel.book.close()

    class Column:
        def __init__(self, char, heading=None, horizontal='left', vertical='bottom', width=None, number_format=None):
            self.char = char.upper()
            self.heading = heading
            self.horizontal = horizontal
            self.vertical = vertical
            self.counter = 1
            self.number_format = number_format

            # if headings are on
            if Excel.headings == True:
                self.counter = 2
                Excel.sheet[f'{self.char}1'] = heading

            # set width
            if width is not None:
                Excel.sheet.column_dimensions[self.char].width = width
            
        def append(self, data):
            cell = f'{self.char}{self.counter}'
            Excel.sheet[cell] = data
            Excel.sheet[cell].alignment = openpyxl.styles.Alignment(horizontal=self.horizontal, vertical=self.vertical)

            if self.number_format is not None:
                Excel.sheet[cell].number_format = self.number_format

            self.counter += 1
        
        def setHeading(self, heading):
            # if headings are on
            if Excel.headings == True:
                self.counter = 2
                Excel.sheet[f'{self.char}1'] = heading