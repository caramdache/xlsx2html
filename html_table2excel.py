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

        self.spans = {}
        self.merged_cells = {}

        self.cell = []
        self.format = {}
        self.mark_color = None
        self.bold = False
        self.italic = False
        self.underline = False
        self.strike = False
        self.li = False
        self.td = False

    def handle_starttag(self, tag, attrs):
        if tag in ['th', 'td']:
            self.td = True
            self.handle_span(attrs)

        elif tag == 'br':
            self.cell.append("\n")

        elif tag == 'li':
            self.li = True

        elif tag in ['b', 'strong']:
            self.bold = True

        elif tag in ['i', 'em', 'blockquote', 'code']:
            self.italic = True

        elif tag == 'u':
            self.underline = True

        elif tag in ['s', 'strike']:
            self.strike = True

        elif tag == 'mark':
            color = 'black'
            for name, value in attrs:
                if name == 'class':
                    color = COLORS[value]
                if name == 'color':
                    color = value

            self.mark_color = color

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
        if self.bold:
            self.format['bold'] = 1

        if self.italic:
            self.format['italic'] = 1

        if self.underline:
            self.format['underline'] = 1

        if self.strike:
            self.format['font_strikeout'] = 1

        if self.mark_color:
            self.format['font_color'] = self.mark_color


        if self.li:
            self.cell.append(f"\n- {data}")

        elif self.td:
            if len(data) > 0:
                self.handle_format()
                self.cell.append(data)

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

        elif tag in ['ul', 'ol']:
            self.cell.append("\n")

        elif tag in ['li']:
            self.li = False

        elif tag in ['b', 'strong']:
            self.bold = False

        elif tag in ['i', 'em', 'blockquote', 'code']:
            self.italic = False

        elif tag == 'u':
            self.underline = False

        elif tag in ['s', 'strike']:
            self.strike = False

        elif tag == 'mark':
            self.mark_color = None

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
        self.format = {}

    def perform_jump(self):
        # Handle successive colspans
        jump = 'jump-before-write'
        while jump == 'jump-before-write':
            rowspan, colspan, jump = self.spans.get(self.col, (None, None, None))

            # Mark this row as processed
            if rowspan is not None:
                self.spans[self.col] = (rowspan - 1, colspan, 'jump-before-write')
                if rowspan == 1:
                    # All colspans have been performed.
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
