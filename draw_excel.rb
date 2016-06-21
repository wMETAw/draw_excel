require 'rubyXL'
require 'RMagick'
include Magick

img = ImageList.new(ARGV[0])

# Excelを作成し、最初のシートを選択
book = RubyXL::Workbook.new
sheet = book[0]

img.each_pixel do |pixel, column, row|

    # ピクセル色を16進数で取得
    color = pixel.to_color(Magick::AllCompliance, false, img.depth, true)
    color.delete!('#')

    # 塗りつぶし
    sheet.add_cell(row, column, '')
    sheet.sheet_data[row][column].change_fill(color)

    # 行、列の幅変更
    sheet.change_column_width(row, 0.01)
    sheet.change_row_height(column, 5)
end

file_name = File.basename("#{ARGV[0]}", '.*')
book.write("#{file_name}.xlsx")