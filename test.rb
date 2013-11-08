require 'rubygems'
require 'spreadsheet'

Spreadsheet.client_encoding = 'UTF-8'

book = Spreadsheet::Workbook.new
sheet1 = book.create_worksheet :name => 'test'

money_format = Spreadsheet::Format.new :number_format => "$#,##0.00"
date_format = Spreadsheet::Format.new :number_format => 'DD.MM.YYYY'

# set default column formats
sheet1.column(1).default_format = money_format
sheet1.column(2).default_format = date_format
sheet1.row(0).push "just text", 5.98, DateTime.now

book.write 'test.xls'