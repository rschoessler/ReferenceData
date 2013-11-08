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

  def insertHeaderRow(numColumns)
    #tmpfile = "new_#{baselineFile}"
    #Spreadsheet.open baselineFile do |book|
      sheet = book.worksheet(0)
      numColumns = sheet.dimensions[3]   #gets the number of columns

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
      #book.write baselineFile
      #return baselineFile
    #end                 #close the spreadsheet

    #File.delete baselineFile                              #need to delete the original file before we can write to it again
    #FileUtils.move tmpfile, baselineFile, :force => true  #move the file
  end

  def generateBaselineFile(baselineFile,numCols,numRows)

    tmpfile = "new_#{baselineFile}"
    Spreadsheet.open baselineFile do |book|
      sheet = book.worksheet(0)

        puts "Number of rows:  #{numRows}"
        puts "Number of columns:  #{numCols}"

        ##create the header row
        #letter = "a"
        #i = 0
        #begin
        #  sheet[0,i] = letter.upcase
        #  i += 1
        #  letter = letter.next
        #  puts letter.upcase
        #  puts i
        #end until i == numCols

        rowIndex=0

        #first go through one row
        begin                           #rows
          row = sheet.row(rowIndex)
          colIndex=0                    #set this back to zero
          begin                         #column
            value = row[colIndex].to_s
            puts "Value:  #{value}"          #replace with code to replace via regex
            format = row.format(colIndex).number_format
            puts "Format:  |#{format}|"
            value = row[colIndex].to_s
            newValue = value.gsub(/\./, '\.')           #escape all '.'
            newValue = newValue.gsub(/\*/, '\*')        #escape all '*'
            newValue = newValue.gsub(/\(/, '\(')        #escape all '('
            newValue = newValue.gsub(/\)/, '\)')        #escape all ')'
            newValue = newValue.gsub(/(\d\d)-(\D\D\D)-(\d\d\d\d)/, '\1-\2-\3')
            newValue = newValue.gsub(/(\d+)\/(\d+)\/(\d\d+)/, '="\1/\2/\3"')
            newValue = newValue.gsub(/^(-?\d+)$/,'\1\.0') #get all values like -85000000 or 85000000 with no decimal
            puts "New Value:  #{newValue}"
            puts "Coordinates: #{rowIndex}, #{colIndex}"
            puts "iterations = #{colIndex}"
            if value != newValue
              sheet[rowIndex,colIndex] = newValue
              puts sheet[rowIndex,colIndex]
            end
            colIndex += 1
          end until colIndex == numCols
          rowIndex += 1
        end until rowIndex == numRows
      book.write tmpfile
    end                 #close the spreadsheet

    #File.delete baselineFile                              #need to delete the original file before we can write to it again
    #FileUtils.move tmpfile, baselineFile, :force => true  #move the file

  end

  def getColumnCount(baselineFile)
    Spreadsheet.open baselineFile do |book|
      sheet = book.worksheet(0)
      #first get the dimensions
      #puts sheet.dimensions
      numRows = sheet.dimensions[1]
      numCols = sheet.dimensions[3]

      puts "Number of rows:  #{numRows}"
      puts "Number of columns:  #{numCols}"
      return numCols, numRows
    end

  end

  def createGenericTestFile(baselineFile)
    tmpfile = "tmp_#{baselineFile}"
    Spreadsheet.open baselineFile do |book|
      sheet = book.worksheet "sheet1"
      sheet[1,0] = "(300,000,000.00)"
      sheet[1,1] = "253.45%"
      sheet[1,2] = "(300,000,000.00)"
      sheet[1,3] = "*fix this asterisk"
      sheet[2,1] = "01-Jan-2013"
      sheet[2,2] = "12/1/2013"
      book.write tmpfile
    end                 #close the spreadsheet

    File.delete baselineFile                              #need to delete the original file before we can write to it again
    FileUtils.move tmpfile, baselineFile, :force => true  #move the file

  end

end


