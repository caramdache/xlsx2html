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
- `<img src="<path">`
- `<br>`

## How to use

### HTML to XLSX

We rely on Python's excellent [XLSXWriter](https://xlsxwriter.readthedocs.io/) to generate the XLSX file.

```
import xlsxwriter
import html2excel.py

workbook = xlsxwriter.Workbook('table.xlsx')
worksheet = workbook.add_worksheet()

html = '... some HTML table...'
p = HTMLTable2Excel(wb, ws, default_format={
    'font_name': 'Arial',
    'font_size': 10,
    'text_wrap': 1,
    'valign': 'top',
    'border': 1,
    'border_color': '#0000ff',
})                                                                
p.feed(html)

workbook.close()
```

### XLSX to HTML

There is no support for rich strings in [openpyxl](https://openpyxl.readthedocs.io/en/stable/) today unfortunately, so we use instead [rubyXL](https://github.com/weshatheleopard/rubyXL) instead.

```
require 'rubyXL'
require 'rubyXL/convenience_methods'
require './excel2html'

wb = RubyXL::Parser.parse('some excel file.xlsx')

wb.worksheets.each { |ws|
    worksheet_to_html(ws)
}
```

## Examples

### Example 1 (rich text)

```
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

```
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

```
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

