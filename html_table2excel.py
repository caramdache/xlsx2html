from html.parser import HTMLParser

class HTMLTable2Excel(HTMLParser):
    def __init__(
        self,
        workbook,
        worksheet,
        decode_html_entities=False,
    ):

        HTMLParser.__init__(self)

        self.workbook = workbook
        self.worksheet = worksheet
        self.row = 0
        self.col = 0

        self.parse_html_entities = decode_html_entities

        self.td = False
        self.cell = []
        self.format = {}

    def handle_starttag(self, tag, attrs):
        if tag in ['th', 'td']:
            self.td = True
        elif tag == 'mark':
            color = attrs[0][1]
            self.format['font_color'] = color
        elif tag == 'b':
            self.format['bold'] = True
        elif tag == 'i':
            self.format['italic'] = True
        elif tag == 'u':
            self.format['underline'] = True
        elif tag == 's':
            self.format['font_strikeout'] = True

    def handle_data(self, data):
        if self.td:
            striped_data = data.strip()
            if len(striped_data) > 0:
                self.handle_format()
                self.cell.append(f"{striped_data} ")

    def handle_format(self):
        if len(self.format) >= 1:
            format = self.workbook.add_format(self.format)
            self.cell.append(format)
            self.format = {}

    def handle_charref(self, name):
        if self.parse_html_entities:
            self.handle_data(self.unescape('&#{};'.format(name)))

    def handle_endtag(self, tag):
        if tag in ['td', 'th']:
            if len(self.cell) > 2:
                self.worksheet.write_rich_string(self.row, self.col, *self.cell)
            elif len(self.cell) == 2:
                self.worksheet.write_string(self.row, self.col, self.cell[1], self.cell[0])
            else:
                self.worksheet.write_string(self.row, self.col, self.cell[0])

            self.cell = []
            self.td = False
            self.col += 1

        elif tag == 'tr':
            self.worksheet.set_column(self.row, self.col, 20)
            self.row += 1
            self.col = 0
