#!/usr/bin/env ruby

require 'rubyXL'
require 'rubyXL/convenience_methods'

require 'securerandom'


MARKERS = {
    'F79646' => 'marker-orange',
    'E46C0A' => 'marker-orange',
    'FFC000' => 'marker-orange',
    'E6B9B8' => 'marker-orange',
    'D99694' => 'marker-orange',
    'FF9900' => 'marker-orange',

    '996633' => 'marker-brown',
    '984807' => 'marker-brown',
    '948A54' => 'marker-brown',
    'CC9900' => 'marker-orange',
    'CC6600' => 'marker-orange',

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
    '4BACC6' => 'marker-blue',
    '558ED5' => 'marker-blue',
    'B7DEE8' => 'marker-blue',
    '93CDDD' => 'marker-blue',

    '31859C' => 'marker-dark-blue',
    '4F81BD' => 'marker-dark-blue',

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
    '9900FF' => 'marker-purple',
    '9933FF' => 'marker-purple',

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


# FIX limitation in RubyXL
module RubyXL
  module ColorConvenienceMethods
    def get_rgb(workbook)
      if rgb then
        return rgb
      elsif theme then
        theme_color = workbook.theme.get_theme_color(theme)
        
        rgb_color = theme_color && theme_color.a_srgb_clr
        color_value = rgb_color && rgb_color.val

        # FIX - Handle system colors
        unless color_value then
          rgb_color = theme_color && theme_color.a_sys_clr
          color_value = rgb_color && rgb_color.last_clr
        end
        # END FIX

        return nil if color_value.nil?

        RubyXL::RgbColor.parse(color_value).to_hls.apply_tint(tint).to_rgb.to_s
      end
    end
  end
end


module RubyXL
  class Worksheet
    def images
        @images
    end

    def images=(array)
        @images = Hash.new{|h,k| h[k]=[]}

        array.each { |img|
            # Fix image format (.wmf is not valid, but .emf is)
            img.format = img.format.gsub(/wmf/i, 'emf')

            from = img.anchor._from
            to = img.anchor.to if img.anchor.kind_of?(TwoCellAnchor)

            # Images in merged ranges will be ignored, except the ones in the top-left corner.
            # So we reassign merged images to the top-left cell.
            self.merged_cells.each { |mcell|
                row_range, col_range = mcell.ref.row_range, mcell.ref.col_range

                if row_range.member?(from.row) && col_range.member?(from.col)
                    unless row_range.first == from.row && col_range.first == from.col
                        # TODO: Use some average, like the barycenter, instead of top-left corner
                        dr = from.row - row_range.first
                        dc = from.col - col_range.first

                        from.row -= dr
                        from.col -= dc

                        to.row -= dr if to
                        to.col -= dc if to
                    end
                end
            }

            @images[[from.row, from.col]].append(img)
        }
    end
  end
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

    if rgb =~ /^(FF)?000000$/ || rgb =~ /0D0D0D/
        nil
    elsif rgb =~ /^(.{2,2})(.{6,6})$/
        $2
    else
        rgb
    end
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

def marker(rgb)
    return nil if rgb.nil?

    marker = MARKERS[rgb.upcase]
    puts ">>> marker missing for: #{rgb}" if marker.nil?
    marker
end

def styles(element)
    {
             b: element && element.b      && (element.b      != false), 
             i: element && element.i      && (element.i      != false), 
             u: element && element.u      && (element.u      != false), 
        strike: element && element.strike && (element.strike != false),
    }
end

def header_to_html()
    """
<table>
    <tbody>
"""
end

def footer_to_html()
    """
    </tbody>
</table>
"""
end

def worksheet_to_html(worksheet)
    s = header_to_html()

    s << rows_to_html(worksheet)

    s << footer_to_html()
end

def rows_to_html(worksheet)
    s = ''

    worksheet.each_with_index { |row, i|
        unless row.nil?
            if row.cells.count > 0
                s << "<tr>\n"

                row.cells.each_with_index { |cell, j|
                    if cell
                        s << cell_to_html(cell) unless omit?(cell)
                    else
                        s << "<td></td>\n"
                    end
                }

                s << "</tr>\n"
            end
        end
    }

    s
end

def cell_to_html(cell)
    s = "<td#{span(cell)}#{fill(cell)}>"

    # puts "Cell(#{cell.row}, #{cell.column})"
    # s << "<span  style='font-size: 8px;'>(#{cell.row}, #{cell.column})</span><br>\n"

    s << value_to_html(cell)

    if defined?(TwoCellAnchor)
        s << image_to_html(cell)
    end

    s << "</td>\n"
end

def value_to_html(cell)
    s = ''

    defaults = styles(cell.get_cell_font)

    if cell.value_container.nil?
        s << ''

    elsif cell.datatype != 's'
        s << cell.value_container.value

    else
        shared_strings = cell.worksheet.workbook.shared_strings_container
        rich_text = shared_strings[cell.raw_value.to_i]

        if rich_text.r.count > 0
            rich_text.r.each { |run|
                s << run_to_html(
                    cell,
                    run.t.value,
                    run.r_pr ? run.r_pr.color : font_color(cell),
                    styles(run.r_pr),
                    defaults
                )
            }
        else
            s << run_to_html(
                cell,
                cell.value,
                font_color(cell),
                {},
                defaults
            )
        end
    end

    s.gsub("\n", "<br>\n").gsub(/\n( )+/) { |match| "\n#{'&nbsp;' * match.length}"}
end

def run_to_html(cell, value, color, locals, defaults)
    s = ''

    s << '<b>' if locals[:b] || defaults[:b]
    s << '<i>' if locals[:i] || defaults[:i]
    s << '<u>' if locals[:u] || defaults[:u]
    s << '<strike>' if locals[:strike] || defaults[:strike]

    if color
        rgb = rgb(color, cell)
        marker = marker(rgb)

        s << "<mark class='#{marker}'>" unless marker.nil?
        s << value
        s << '</mark>' unless marker.nil?
    else
        s << value
    end

    s << '</b>' if locals[:b] || defaults[:b]
    s << '</i>' if locals[:i] || defaults[:i]
    s << '</u>' if locals[:u] || defaults[:u]
    s << '</strike>' if locals[:strike] || defaults[:strike]

    locals.each{|k, v| defaults.delete(k) unless v.nil? }

    s
end

def image_to_html(cell)
    s = ''

    ws = cell.worksheet
    ws.images[[cell.row, cell.column]].each_with_index { |img, i|
        path = '.'
        basename = SecureRandom.uuid #img._id
        ext = img.format

        File.open("#{path}/#{basename}.#{ext}", 'wb') { |f|
            f.write(img.ref.getvalue())
        }

        if ext == 'emf'
            `/usr/bin/inkscape -z --export-plain-svg=#{basename}.svg --file #{basename}.#{ext} && rm #{basename}.#{ext}`
            ext = 'svg'
        end

        s << "<img src='#{path}/#{basename}.#{ext}' style='width:300px;'>"
    }

    s
end

def merged?(cell)
    cell.worksheet.merged_cells.each { |mcell|
        return true if mcell.ref.row_range.member?(cell.row) && mcell.ref.col_range.member?(cell.column)
    }

    false
end

def omit?(cell)
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
