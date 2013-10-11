require 'spreadsheet'
require 'fileutils'
Spreadsheet.client_encoding = 'UTF-8'

#this is the engine that generates the file
#1.  Create a header row
#2.  Replace all . with \.
#3.  Replace all ( with \(
#4.  Replace all ) with \)
#5.  Find all date patterns like 30-Jun-2012 and replace with ="30-Jun-2012"
#   -->> \d\d-\D\D\D-\d\d\d\d
#   -->>  (\d\d)-(\D\D\D)-(\d\d\d\d)
#   -->>  ="\1-\2-\3"

class Baseline

  def initialize
    #do nothing
  end

  def initBaselineFile(numColumns,fileNameRoot)
    #this will create the initial file
    #and takes a parameter of columns to create
    #it will then create a row with A B C etc.

    baselineFile = "#{fileNameRoot}.xls"
    puts baselineFile
    puts numColumns

    #create the file
    book = Spreadsheet::Workbook.new
    sheet = book.create_worksheet
    sheet.name = "sheet1"

    #create the header row
    letter = "a"
    i = 0
    begin
      sheet[0,i] = letter.upcase
      i += 1
      letter = letter.next
      puts letter.upcase
      puts i
    end until i == numColumns

    book.write baselineFile

    return baselineFile

  end

  def replacePeriods(baselineFile)
    tmpfile = "tmp_#{baselineFile}"
    Spreadsheet.open baselineFile do |book|
      sheet = book.worksheet "sheet1"
      sheet[1,0] = "."
      sheet[1,1] = "."

      book.write tmpfile
    end                 #close the spreadsheet

    File.delete baselineFile                              #need to delete the original file before we can write to it again
    FileUtils.move tmpfile, baselineFile, :force => true  #move the file

  end

end


