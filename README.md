# HTML Table to Excel (and vice-versa)

Export an HTLM table to Excel, or export an Excel table to HTML.

The following features are supported:

- `<table>`
- `<td rowspan="2" colspan="3>`
- `<b>` (bold)
- `<i>` (italics)
- `<u>` (underline)
- `<s>` (strikethrough)
- `<mark color="red">` (mark)

## HTML table to .xlsx

```
import xlsxwriter
import html_table2excel.py

workbook = xlsxwriter.Workbook('table.xlsx')
worksheet = workbook.add_worksheet()

html = '... some HTML table...'
p = HTMLTable2Excel(workbook, worksheet)                                                                
p.feed(html)

workbook.close()
```

## .xlsx to HTML table

```
require 'rubyXL'
require 'rubyXL/convenience_methods'
require './excel2html_table'

wb = RubyXL::Parser.parse('some excel file.xlsx')

wb.worksheets.each { |ws|
    worksheet_to_html(ws)
}
```

## Example 1 (rich text)

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

## Example 2 (rowspan and colspan)

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

## Example 3

```
<table class='table table-bordered table-hover table-striped'>
  <tr>
    <th ><p>Main</p></td>
    <th ><p>Claim</p></td>
    <th ><p>Text</p></td>
  </tr>
  <tr>
    <td colspan='1' rowspan='4'><p>Main</p></td>
    <td colspan='1' rowspan='4'><p>1</p></td>
    <td ><p>Essai de texte</p></td>
  </tr>
  <tr>
    <td ><p><b>3.2.1 </b><b><u>Section</u></b></p><p>Essai <mark color='#FF0000'>numéro</mark> <mark color='#00B0F0'>deux</mark></p></td>
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

![Alt text](example3.png?raw=true "Example 3")

## Credits

This was inspired by https://github.com/schmijos/html-table-parser-python3
