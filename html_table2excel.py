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
        self.spans = {}
        self.merged_cells = {}
        self.cell = []
        self.format = {}

    def handle_starttag(self, tag, attrs):
        if tag in ['th', 'td']:
            self.td = True
            self.handle_span(attrs)
        elif tag == 'mark':
            color = 'black'
            for name, value in attrs:
                if name == 'class':
                    color = COLORS[value]
                if name == 'color':
                    color = value
            self.format['font_color'] = color
        elif tag == 'b':
            self.format['bold'] = 1
        elif tag == 'i':
            self.format['italic'] = 1
        elif tag == 'u':
            self.format['underline'] = 1
        elif tag == 's':
            self.format['font_strikeout'] = 1

    def handle_span(self, attrs):
        rowspan = colspan = None
        for name, value in attrs:
            rowspan = int(value) if name == 'rowspan' else rowspan
            colspan = int(value) if name == 'colspan' else colspan

        if rowspan:
            colspan = colspan or 1
        if colspan:
            rowspan = rowspan or 1

        if rowspan or colspan:
            self.spans[self.col] = (rowspan, colspan, 'jump-after-write')

    def handle_data(self, data):
        if self.td:
            striped_data = data.strip()
            if len(striped_data) > 0:
                self.handle_format()
                self.cell.append(f"{striped_data} ")

    def handle_format(self):
        # Center merged cells
        # if self.is_merged():
        #     self.format['align'] = 'center'
        #     self.format['valign'] = 'vcenter'
        
        if len(self.format) >= 1:
            format = self.workbook.add_format(self.format)
            self.cell.append(format)

        self.format = {}

    def handle_charref(self, name):
        if self.parse_html_entities:
            self.handle_data(self.unescape(f"&#{name};"))

    def handle_endtag(self, tag):
        if tag == 'tr':
            self.handle_tr()
        elif tag in ['td', 'th']:
            self.handle_td()

    def handle_tr(self):
        # Handle colspans followed immediately by </tr>.
        rowspan, colspan, jump = self.perform_jump()

        self.row += 1
        self.col = 0

    def handle_td(self):
        # Handle consecutive colspans.
        rowspan, colspan, jump = self.perform_jump()

        if jump == 'jump-after-write':
            # Process the cell that starts the rowspan/colspan.
            self.worksheet.merge_range(
                self.row, self.col,
                self.row + rowspan - 1, self.col + colspan - 1,
                ''
            )
            self.merged_cells[(self.row, self.col)] = True

            self.write_cell()
            self.col += colspan
        else:
            # Process non rowspan/colspan cells
            self.write_cell()
            self.col += 1

        self.td = False

    def perform_jump(self):
        # Handle successive colspans
        jump = 'jump-before-write'
        while jump == 'jump-before-write':
            rowspan, colspan, jump = self.spans.get(self.col, (None, None, None))

            # Mark this row as processed
            if rowspan is not None:
                self.spans[self.col] = (rowspan - 1, colspan, 'jump-before-write')
                if rowspan == 1:
                    # All colspans have been processed.
                    del self.spans[self.col]

            # Skip colspan columns
            if jump == 'jump-before-write':
                self.col += colspan

        return (rowspan, colspan, jump)

    def write_cell(self):
        # Prepare to handle display of long strings.
        wrap = self.workbook.add_format({'text_wrap': 1, 'valign': 'top'})

        count = len(self.cell)
        if count > 2:
            self.cell.append(wrap)
            res = self.worksheet.write_rich_string(self.row, self.col, *self.cell)

        elif count == 2:
            # Handle the case of 2 strings in a row. Work around write_rich_string.
            if type(self.cell[0]) == str:
                self.cell = [wrap, f"{self.cell[0]}\n{self.cell[1]}"]

            format, data = self.cell
            format.set_text_wrap()
            format.set_align('top')
            res = self.worksheet.write_string(self.row, self.col, data, format)

        elif count == 1:
            wrap = self.workbook.add_format({'text_wrap': 1, 'valign': 'top'})
            res = self.worksheet.write_string(self.row, self.col, self.cell[0], wrap)

        else:
            # Tag was empy, no action.
            res = 0

        if res < 0:
            print(f"{res}: {self.cell}\n")

        self.cell = []

    def is_merged(self):
        return (self.row, self.col) in self.merged_cells
