# html_table2excel
Export an HTLM table to Excel.

I've been wanting something simple to export an HTML with formatting to Excel and here is what I have come up with.

## Example

Given the following table:

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

the following code will save the table to an Excel spreadsheet:

```

workbook = xlsxwriter.Workbook('rich_strings.xlsx')
worksheet = workbook.add_worksheet()

p = HTMLTable2Excel(workbook, worksheet)                                                                
p.feed(html_string)

workbook.close()
```

## Credits

This was inspired by https://github.com/schmijos/html-table-parser-python3
