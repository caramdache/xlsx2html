from html.parser import HTMLParser
import xlsxwriter

class HTMLTable2Excel(HTMLParser):
    def __init__(
        self,
        workbook,
        worksheet,
        decode_html_entities=False,
    ):

        HTMLParser.__init__(self)

        self._workbook = workbook
        self._worksheet = worksheet
        self._row = 0
        self._col = 0

        self._parse_html_entities = decode_html_entities

        self._in_td = False
        self._in_th = False
        self._current_cell = []
        self._current_format = {}
        self.tables = []

    def handle_starttag(self, tag, attrs):
        if tag == 'td':
            self._in_td = True
        elif tag == 'th':
            self._in_th = True
        elif tag == 'mark':
            color = attrs[0][1]
            self._current_format['font_color'] = color
        elif tag == 'b':
            self._current_format['bold'] = True
        elif tag == 'i':
            self._current_format['italic'] = True
        elif tag == 'u':
            self._current_format['underline'] = True
        elif tag == 's':
            self._current_format['font_strikeout'] = True

    def handle_data(self, data):
        if self._in_td or self._in_th:
            striped_data = data.strip()
            if len(striped_data) > 0:
                self.handle_format()
                self._current_cell.append(f"{striped_data} ")

    def handle_format(self):
        if len(self._current_format) >= 1:
            format = self._workbook.add_format(self._current_format)
            self._current_cell.append(format)
            self._current_format = {}

    def handle_charref(self, name):
        if self._parse_html_entities:
            self.handle_data(self.unescape('&#{};'.format(name)))

    def handle_endtag(self, tag):
        if tag == 'td':
            self._in_td = False
        if tag == 'th':
            self._in_th = False

        if tag in ['td', 'th']:
            if len(self._current_cell) > 1:
                self._worksheet.write_rich_string(self._row, self._col, *self._current_cell)
            else:
                self._worksheet.write_string(self._row, self._col, self._current_cell[0])

            self._current_cell = []
            self._col += 1
        elif tag == 'tr':
            self._worksheet.set_column(self._row, self._col, 20)
            self._row += 1
            self._col = 0
