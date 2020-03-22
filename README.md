# html_table2excel

Export an HTLM table to Excel. Supports the following features:

- `<table>`
- `<td rowspan="2" colspan="3>`
- `<b>` (bold)
- `<i>` (italics)
- `<u>` (underline)
- `<s>` (strikethrough)
- `<mark color="red">` (mark)

## How to use

```
import xlsxwriter

workbook = xlsxwriter.Workbook('table.xlsx')
worksheet = workbook.add_worksheet()

html = '... some HTML table...'
p = HTMLTable2Excel(workbook, worksheet)                                                                
p.feed(html)

workbook.close()
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

## Credits

This was inspired by https://github.com/schmijos/html-table-parser-python3
