COLORS = {
    'marker-yellow':     '#fdfd77',
    'marker-green':      '#63f963',
    'marker-pink':       '#fc7999',
    'marker-blue':       '#72cdfd',

    'marker-orange':     '#FFC000',
    'marker-purple':     '#c6a0ff',
    'marker-brown':      '#bc9a79',
    'marker-terra-cota': '#999787',
    'marker-brique':     '#cea1a1',
    'marker-red':        '#ff4040',
    'marker-dark-blue':  '#8199b6',
    'marker-grey':       '#d9d9d9',
}


class HTMLTable2Excel(HTMLParser):
    def __init__(
        self,
        workbook,
        worksheet,
        default_format={},
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
        self.default_format = default_format
        self.format = {}
        self.font_size = None
        self.text_align = None
        self.fill_color = None
        self.mark_color = None
        self.bold = False
        self.italic = False
        self.underline = False
        self.strike = False
        self.li = False
        self.td = False
        self.skip = False

    def set_font(self):
        self.format['font_name'] = self.default_format.get('font_name', 'Arial')
        self.format['font_size'] = self.default_format.get('font_size', 10)

        if self.font_size is not None and self.font_size != 'text-default':
            self.format['font_size'] = 20

    def handle_starttag(self, tag, attrs):
        if tag in ['th', 'td']:
            self.td = True
            self.handle_skip(attrs)
            self.handle_colspan(attrs)
            self.text_align = self.get_style_attr(attrs, 'text-align')
            self.fill_color = self.get_style_attr(attrs, 'background-color')

        elif tag == 'br':
            self.set_font()
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
            self.handle_mark(attrs)

        elif tag == 'span':
            self.handle_span(attrs)

    def handle_span(self, attrs):
        for name, value in attrs:
            if name == 'class' and 'text-' in value:
                self.font_size = value

        self.mark_color = self.get_style_attr(attrs, 'color')

    RE_CCS_SELECTORS = re.compile(r'([^:;\s]+)\s?:\s?([^;\s]+)(?=;)?') 

    def get_style_attr(self, attrs, attr):
        for name, value in attrs:
            if name == 'style':
                for attribute, content in re.findall(self.RE_CCS_SELECTORS, value):
                    if attribute == attr:
                        return content
        return None

    def handle_mark(self, attrs):
        color = 'black'
        for name, value in attrs:
            if name == 'class':
                color = COLORS[value]
            if name == 'color':
                color = value

        self.mark_color = color

    def handle_skip(self, attrs):
        for name, value in attrs:
            if name == 'class' and value == 'skip':
                self.skip = True

    def handle_colspan(self, attrs):
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
        data = data.strip()

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
            self.handle_format()
            self.cell.append(f"\n- {data}")

        elif self.td: 
            if len(data) > 0:
                self.handle_format()
                self.cell.append(data)

            self.format = {}
            # self.font_size = None

    def handle_format(self):
        self.set_font()
        format = self.workbook.add_format(self.format)
        self.cell.append(format)

    def handle_charref(self, name):
        if self.parse_html_entities:
            self.handle_data(self.unescape(f"&#{name};"))

    def handle_endtag(self, tag):
        if tag == 'tr':
            self.handle_end_tr()

        elif tag in ['td', 'th']:
            self.handle_end_td()

        elif tag in ['ul', 'ol']:
            self.set_font()
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

        elif tag == 'span':
            self.font_size = None
            self.mark_color = None

    def handle_end_tr(self):
        # Handle colspans followed immediately by </tr>.
        self.perform_jump()

        self.row += 1
        self.col = 0

    def handle_end_td(self):
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
            
            if not self.skip:
                self.write_cell()
            self.col += colspan
        else:
            # Process other cells
            if not self.skip:
                self.write_cell()
            self.col += 1

        self.td = False
        self.skip = False
        self.cell = []
        self.format = {}
        self.font_size = None
        self.text_align = None
        self.fill_color = None

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
        cell_format = self.workbook.add_format(self.default_format)

        if self.text_align:
            cell_format.set_align(self.text_align)
            if self.text_align == 'center':
                cell_format.set_align('vcenter')

        if self.fill_color:
            cell_format.set_bg_color(self.fill_color)

        count = len(self.cell)
        if count >= 2:
            if count == 2:
                # Work around write_rich_string's limitation. Add an invisible zero-width-space
                self.cell = ['\u200b'] + self.cell

            self.cell.append(cell_format)
            res = self.worksheet.write_rich_string(self.row, self.col, *self.cell)

        elif count == 1:
            res = self.worksheet.write_string(self.row, self.col, self.cell[0], cell_format)

        else:
            res = self.worksheet.write_blank(self.row, self.col, '', cell_format)

        if res < 0:
            print(f"{res}: {self.cell}\n")

    def is_merged(self):
        return (self.row, self.col) in self.merged_cells
