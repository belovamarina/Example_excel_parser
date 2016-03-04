require 'rubygems'
require 'watir-webdriver'
require 'spreadsheet'
require 'roo'
require 'rubyXL'

# Column name with codes we need
COL_NAME = 'Код ресурса по Реестру цен'

# Open file
xls = Roo::Spreadsheet.open('excel.xlsx')

# Set working sheet
xls.default_sheet = xls.sheets.first

# Set begining of the table  (key: header name, value: cell value)
hashes_array = xls.parse(header_search: [COL_NAME])

# Save codes to array
@array = []
hashes_array.each do |hash|
  if hash[COL_NAME].is_a?(Numeric) && hash[COL_NAME].to_s.size >= 7
    @array << hash[COL_NAME]
  end
end

xls.close

print @array.uniq # > [3311605, 2522119, 3697492, 3311615]

# Create directory for data

DATA_DIR = 'okdp_data'
Dir.mkdir(DATA_DIR) unless File.exist?(DATA_DIR)

BASE_URL = 'http://cmec.spb.ru/dm5/index/reestr'

def save_table(table, filename)
  if table.exists?
    File.open(filename, 'a') do |fp|
      table.trs.each do |tr|
        fp.puts tr.tds.map(&:text).join('| ')
      end
    end
  else
    puts 'Table not found'
  end
end

# Save tables from base url to directory

browser = Watir::Browser.new

@array.uniq.each do |okdp|
  @okdp = okdp.to_s
  @local_fname = "#{DATA_DIR}/#{File.basename(@okdp)}.txt"

  browser.goto BASE_URL

  browser.text_field(name: 'code').set @okdp

  browser.link(id: 'searchBut').click

  @table_array = []
  table = browser.table(index: 2)

  save_table(table, @local_fname)

  if browser.ol(class: 'paging').exists?
    links_count = browser.ol(class: 'paging').links.length
    (2..(links_count + 1)).each do |i|
      browser.link(text: i.to_s).click
      save_table(table, @local_fname)
      sleep rand(3..7)
    end
  end
end
