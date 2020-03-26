require 'rubyXL'
require 'rubyXL/convenience_methods'

MARKERS = {
    'F79646' => 'marker-orange',
    'E46C0A' => 'marker-orange',
    'FFC000' => 'marker-orange',
    'E6B9B8' => 'marker-orange',
    'D99694' => 'marker-orange',

    '996633' => 'marker-brown',
    '984807' => 'marker-brown',
    '948A54' => 'marker-brown',

    '4A452A' => 'marker-terra-cota',

    'FFFF00' => 'marker-yellow',

    '008080' => 'marker-green',
    '006600' => 'marker-green',
    '009900' => 'marker-green',
    '00B050' => 'marker-green',
    '92D050' => 'marker-green',
    '9BBB59' => 'marker-green',
    '77933C' => 'marker-green',
    '4F6228' => 'marker-green',

    '0000FF' => 'marker-blue',
    '0070C0' => 'marker-blue',
    '00B0F0' => 'marker-blue',
    '31859C' => 'marker-blue',
    '4BACC6' => 'marker-blue',
    '4F81BD' => 'marker-blue',
    '558ED5' => 'marker-blue',
    'B7DEE8' => 'marker-blue',
    '93CDDD' => 'marker-blue',

    '1F497D' => 'marker-dark-blue',
    '376092' => 'marker-dark-blue',
    '002060' => 'marker-dark-blue',
    '10253F' => 'marker-dark-blue',
    '17375E' => 'marker-dark-blue',
    '215968' => 'marker-dark-blue',
    '254061' => 'marker-dark-blue',

    '6600FF' => 'marker-purple',
    '7030A0' => 'marker-purple',
    '8064A2' => 'marker-purple',
    'B3A2C7' => 'marker-purple',
    'CCC1DA' => 'marker-purple',
    '604A7B' => 'marker-purple',

    'FF66CC' => 'marker-pink',
    'FF00FF' => 'marker-pink',

    'C00000' => 'marker-red',
    'FF0000' => 'marker-red',
    'C0504D' => 'marker-red',

    '953735' => 'marker-brique',
    '632523' => 'marker-brique',

    '808080' => 'marker-grey',
    'A6A6A6' => 'marker-grey',
    'BFBFBF' => 'marker-grey',
    'D9D9D9' => 'marker-grey',
}


def marker(rgb)
    return nil if rgb.nil?

    marker = MARKERS[rgb.upcase]
    print(">>> marker missing for: #{rgb}\n") if marker.nil?
    marker
end

def font_color(cell)
    # Instead of convenience method: cell.font_color, which is buggy
    cell.get_cell_font.color
end

def rgb(color, cell)
    return nil if color.nil?

    rgb = if color.is_a? String
        color
    else
        color.get_rgb(cell.worksheet.workbook)
    end

    rgb = rgb.upcase unless rgb.nil?

    if rgb =~ /^(FF)?000000$/
        nil
    elsif rgb =~ /^(.{2,2})(.{6,6})$/
        $2
    else
        rgb
    end
end

def worksheet_to_html(worksheet)
    s = """<table>
    <tbody>
"""
    s << rows_to_html(worksheet)

    s << """
    </tbody>
</table>
"""
end

def rows_to_html(worksheet)
    s = ''

    worksheet.each_with_index { |row, i|
        unless row.nil?
            s << "<tr>"

            row.cells.each_with_index { |cell, j|
                s << cell_to_html(cell, i, j) unless omit?(cell)
            }

            s << "</tr>"
        end
    }

    s
end

def cell_to_html(cell, i, j)
    s = "<td#{span(cell)}#{fill(cell)}>"

    # print("Cell(#{cell.row}, #{cell.column})\n")
    # s << "<span  style='font-size: 8px;'>(#{cell.row}, #{cell.column})</span><br>\n"

    s << value_to_html(cell)

    s << '</td>'
end

def fill(cell)
    color = cell.fill_color

    rgb = rgb(color, cell)
    if rgb && rgb !~ /FFFFFF/
        " style='background-color:##{rgb};'"
    else
        ''
    end
end

def value_to_html(cell)
    s = ''

    if cell.value_container.nil?
        s << ''

    elsif cell.datatype != 's'
        s << cell.value_container.value

    else
        shared_strings = cell.worksheet.workbook.shared_strings_container
        rich_text = shared_strings[cell.raw_value.to_i]

        if rich_text.r.count > 0
            font = cell.get_cell_font
            defaults = {b: font.b, i: font.i, u: font.u, strike: font.strike}

            rich_text.r.each { |run| s << run_to_html(cell, run, defaults) }
        else
            rgb = rgb(font_color(cell), cell)
            marker = marker(rgb)

            s << "<mark class='#{marker}'>" unless marker.nil?
            s << cell.value
            s << '</mark>' unless marker.nil?
        end
    end

    s = s.gsub("\n", "<br>")

    s
end

def run_to_html(cell, run, defaults)
    s = ''

    pr = run.r_pr

    locals = {b: pr && pr.b, i: pr && pr.i, u: pr && pr.u, strike: pr && pr.strike}

    s << '<b>' if locals[:b] || defaults[:b]
    s << '<i>' if locals[:i] || defaults[:i]
    s << '<u>' if locals[:u] || defaults[:u]
    s << '<strike>' if locals[:strike] || defaults[:strike]

    if pr
        rgb = rgb(pr.color, cell)
        marker = marker(rgb)
        s << "<mark class='#{marker}'>" unless marker.nil?
        s << run.t.value
        s << '</mark>' unless marker.nil?
    else
        s << run.t.value
    end

    s << '</b>' if locals[:b] || defaults[:b]
    s << '</i>' if locals[:i] || defaults[:i]
    s << '</u>' if locals[:u] || defaults[:u]
    s << '</strike>' if locals[:strike] || defaults[:strike]

    locals.each{|k, v| defaults.delete(k) unless v.nil? }

    s
end

def merged?(cell)
    cell.worksheet.merged_cells.each { |mcell|
        return true if mcell.ref.row_range.member?(cell.row) && mcell.ref.col_range.member?(cell.column)
    }

    false
end

def omit?(cell)
    return true if cell.nil?

    cell.worksheet.merged_cells.each { |mcell|
        if mcell.ref.row_range.member?(cell.row) && mcell.ref.col_range.member?(cell.column)
            if mcell.ref.row_range.first == cell.row && mcell.ref.col_range.first == cell.column
                return false
            else
                return true
            end
        end
    }

    false
end

def span(cell)
    cell.worksheet.merged_cells.each { |mcell|
        if mcell.ref.row_range.first == cell.row && mcell.ref.col_range.first == cell.column
            return " colspan='#{mcell.ref.col_range.size}' rowspan='#{mcell.ref.row_range.size}'"
        end
    }

    ''
end
