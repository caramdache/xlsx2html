# Excel to HTML (and back!)

Export an HTLM table to Excel, or export an Excel table to HTML.

## Features at a glance

- `<table>`
- `<td rowspan="2" colspan="3">`
- `<td style="background-color:red; text-align:center;">`
- `<td class="skip">`
- `<li>`, `<ul>` and `<ol>`
- `<b>` and `<strong>` 
- `<i>`, `<em>`, `<blockquote>` and `<code>`
- `<u>` 
- `<s>` and `<strike>`
- `<mark color="red">`
- `<span style="color:red" class="text-big">` 
- `<img src="some_path">`
- `<br>`

## How to use

### HTML to XLSX

We rely on Python's excellent [XLSXWriter](https://xlsxwriter.readthedocs.io/) to generate the XLSX file.

```python
#!/usr/bin/env python3

import xlsxwriter

with open('html2excel.py') as infile:
    exec(infile.read())

wb = xlsxwriter.Workbook('test.xlsx')
ws = wb.add_worksheet()

p = HTML2Excel(wb, ws, default_format={
    'font_name': 'Arial',
    'font_size': 10,
    'text_wrap': 1,
    'valign': 'top',
    'border': 1,
    'border_color': '#0000ff',
})    

with open('test.html') as input:
    html = input.read()
    p.feed(html)

wb.close()
```

#### Image scaling

If you text-wrap cells or if you merge cells, your images may be squeezed. In that case, you may use the following workaround:

```python
# ... (same as above)

with open('test.html') as input:
    html = input.read()
    image_paths = p.feed(html)

wb.close()

# Now switch to a different library to add images with no squeeze

import openpyxl
from xlsxwriter.utility import xl_rowcol_to_cell

wb = openpyxl.load_workbook('test.xlsx')
ws = wb.active

for row, col, path in image_paths:
    image = openpyxl.drawing.image.Image(path)
    ws.add_image(image, xl_rowcol_to_cell(row, col))

wb.save('test.xlsx')
```

### XLSX to HTML

There is no support for rich strings in [openpyxl](https://openpyxl.readthedocs.io/en/stable/), so we use [rubyXL](https://github.com/weshatheleopard/rubyXL). They are both excellent libraries.

```ruby
#!/usr/bin/env ruby

require 'rubyXL'
require 'rubyXL/convenience_methods'

require './excel2html'

wb = RubyXL::Parser.parse('some excel file.xlsx')

wb.worksheets.each { |ws|
    worksheet_to_html(ws)
}
```

#### Also export images

There is a little bit of a conendrum:

- openpyxl does not support rich text, so we use RubyXL; however
- rubyxl does not support images, so we also need to use openpyxl

Fortunately, `pycall` comes to the rescue and allows us to use Python code inside Ruby.

```ruby
#!/usr/bin/env ruby

require 'rubyXL'
require 'rubyXL/convenience_methods'

require 'pycall/import'
include PyCall::Import
pyimport :openpyxl
pyfrom 'openpyxl.drawing.spreadsheet_drawing', import: 'TwoCellAnchor'

require './excel2html'

wb = RubyXL::Parser.parse('test.xlsx')
wb2 = openpyxl.load_workbook('test.xlsx')

wb.worksheets.each_with_index { |ws, i|
    # Index images by cell row/col for easier later retrieval
    ws.images = wb2.worksheets[i]._images
    
    puts ws.sheet_name
    html = worksheet_to_html(ws)

    File.open("test_#{ws.sheet_name}.html", 'wb') { |f|
        f.write(html)
    }
}
``` 

## Examples

### Example 1 (rich text)

```html
<table>
    <thead>
        <tr>
            <th>C1</th><th>C2</th><th>C3</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>Basic string</td>
            <td><mark class="red"><u>A string</u></mark></td>
            <td><u><mark class="red"><i>Another</i></mark></u> bizarre <u>string</u></td>
            <td>A <mark class="red">third</mark> <mark class="blue">enourmous</mark> string</td>
        </tr>
    </tbody>
</table>
```

![Alt text](example1.png?raw=true "Example 1")

### Example 2 (rowspan and colspan)

```html
<table>
    <tbody>
        <tr><td>a</td><td>b</td><td>c</td><td>d</td><td>e</td></tr>
        <tr><td>a</td><td rowspan="3" colspan="3">A <mark class="red">third</mark> <mark class="blue">enourmous</mark> string</td><td>e</td></tr>
        <tr><td>a</td><td>e</td></tr>
        <tr><td>a</td><td>e</td></tr>
        <tr><td>a</td><td>b</td><td>c</td><td>d</td><td>e</td></tr>
    </tbody>
</table>
```

![Alt text](example2.png?raw=true "Example 2")

### Example 3

```html
<table class='table table-bordered table-hover table-striped'>
  <tr>
    <th ><p>Col1</p></td>
    <th ><p>Col2</p></td>
    <th ><p>Col3</p></td>
  </tr>
  <tr>
    <td colspan='1' rowspan='4'><p>Merged</p></td>
    <td colspan='1' rowspan='4'><p>1</p></td>
    <td ><p>Some text</p></td>
  </tr>
  <tr>
    <td ><p><b>3.2.1 </b><b><u>Section</u></b></p><p>number <mark color='#FF0000'>two</mark> <mark color='#00B0F0'>three</mark></p></td>
  </tr>
  <tr>
    <td ><p>*nil*</p></td>
  </tr>
  <tr>
    <td ><p>*nil*</p></td>
  </tr>
  <tr>
    <td ><p>11</p></td>
    <td ><p>13</p></td>
  </tr>
</table>
```

![Alt text](example4.png?raw=true "Example 3")

