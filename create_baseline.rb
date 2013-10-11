
require 'spreadsheet'
require_relative 'generate_baseline'
Spreadsheet.client_encoding = 'UTF-8'

mybaseline = Baseline.new

numcols = 8 #abcdefgh
filenameroot = "baseline"


baselineFile = mybaseline.initBaselineFile(numcols, filenameroot)

puts baselineFile

mybaseline.replacePeriods (baselineFile)