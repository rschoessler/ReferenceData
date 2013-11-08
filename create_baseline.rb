require 'spreadsheet'
require_relative 'generate_baseline'
Spreadsheet.client_encoding = 'UTF-8'

mybaseline = Baseline.new

numcols = 4 #abcdefgh etc
filenameroot = "baseline_test"
#baselineFile = "baseline_test.xls"
#baselineFile = "example.xls"
baselineFile = "Valuation_Functions.xls"


#open the baselineFile and count the columns
arrColRow = mybaseline.getColumnCount(baselineFile)
#puts arrColRow
numCols = arrColRow[0]
numRows = arrColRow[1]
puts numCols
puts numRows

#create a file with the header row
#we'll open this up and write the data to it later
#initBaselineFile = mybaseline.initBaselineFile(numcols, filenameroot) #remove this later after testing

#mybaseline.createGenericTestFile (baselineFile)

#puts initBaselineFile

mybaseline.generateBaselineFile(baselineFile,numCols,numRows)

#mybaseline.insertHeaderRow(numCols, baselineFile)

