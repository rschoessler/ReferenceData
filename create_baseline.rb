
require 'spreadsheet'
require_relative 'generate_baseline'
Spreadsheet.client_encoding = 'UTF-8'

mybaseline = Baseline.new

numcols = 4 #abcdefgh etc
filenameroot = "baseline_test"
baselineFile = "baseline_test.xls"


#baselineFile = mybaseline.initBaselineFile(numcols, filenameroot) #remove this later after testing

#mybaseline.replacePeriods (baselineFile)

puts baselineFile

#mybaseline.rewriteBaseline(baselineFile,4)

mybaseline.generateBaselineFile(baselineFile)

#mybaseline.insertHeaderRow(numcols, baselineFile)

