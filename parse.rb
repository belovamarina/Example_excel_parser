require 'rubyXL'

DATA_DIR = 'okdp_data'

dirs = Dir.foreach(DATA_DIR).select { |dir| dir != '.' }.select { |dir| dir != '..' }

dirs.each do |filename|
  puts "Open file #{filename}..."
  file = File.open("#{DATA_DIR}/#{filename}", 'r')
  @code = File.basename("#{DATA_DIR}/#{filename}", '.*')
  strings = file.readlines

  rows_array = []
  strings.each do |line|
    rows_array << line.strip.split('| ')
  end

  # print rows_array
  @hash_array = []
  @garbage_array = []

  rows_array.each do |array|
    if array.length == 8
      h = { id: array[1], code: array[2], name: array[3], unit: array[4], price_net: array[6], price_gross: array[7] }
      @hash_array << h

    elsif array.length == 3
      if array[2].to_f < @hash_array.last[:price_gross].to_f
        @hash_array.last[:price_net] = array[1]
        @hash_array.last[:price_gross] = array[2]
      end
    else
      @garbage_array << array
    end
  end

  print @hash_array

  puts @hash_array.size
  print @garbage_array # array for debug

  workbook = RubyXL::Parser.parse('excel.xlsx')
  worksheet = workbook.worksheets[0]

  puts 'Parse...'

  worksheet.each do |row|
    row && row.cells.each do |cell|
      if cell && cell.value == @code.to_f
        puts "Find Cell with code #{@code}(Row: #{cell.row} Col: #{cell.column})"
        1.upto(@hash_array.size) do |i|
          print "Insert row #{cell.row + i}..."
          worksheet.insert_row(cell.row + i)
        end
      end
    end
  end

  workbook.write('excel.xlsx')

  worksheet.each do |row|
    row && row.cells.each do |cell|
      if cell && cell.value == @code.to_f
        puts "Find Cell with code #{@code}(Row: #{cell.row} Col: #{cell.column})"
        1.upto(@hash_array.size) do |i|
          print 'Change content of cells..'
          worksheet[cell.row + i][cell.column - 1].change_contents(worksheet[cell.row][cell.column - 1].value)
          worksheet[cell.row + i][cell.column].change_contents(@hash_array[i - 1][:code])
          worksheet[cell.row + i][cell.column + 1].change_contents(worksheet[cell.row][cell.column + 1].value)
          worksheet[cell.row + i][cell.column + 2].change_contents(@hash_array[i - 1][:name])
          worksheet[cell.row + i][cell.column + 3].change_contents(@hash_array[i - 1][:unit])
          worksheet[cell.row + i][cell.column + 4].change_contents(worksheet[cell.row][cell.column + 4].value.to_i)
          worksheet[cell.row + i][cell.column + 5].change_contents(@hash_array[i - 1][:price_net].to_f)
          worksheet[cell.row + i][cell.column + 6].change_contents(@hash_array[i - 1][:price_gross].to_f)
          worksheet[cell.row + i][cell.column + 7].change_contents('', "#{worksheet[cell.row + i][cell.column + 4].r}*#{worksheet[cell.row + i][cell.column + 6].r}")
        end
      end
    end
  end
  workbook.write('excel.xlsx')
end

workbook = RubyXL::Parser.parse('excel.xlsx')
worksheet = workbook.worksheets[0]

worksheet.each do |row|
  row && row.cells.each do |cell|
    if cell && cell.column == 7 && cell.value.is_a?(Float)
      puts "Find Cell in column 7 and with numeric value in row: #{cell.row}"
      worksheet[cell.row][8].change_contents('', "#{worksheet[cell.row][5].r} * #{worksheet[cell.row][7].r}")
    elsif cell && cell.column == 4 && cell.value == 'Шт.'
      worksheet[cell.row][8].change_contents(0.00)
    end
  end
end

workbook.write('excel.xlsx')
