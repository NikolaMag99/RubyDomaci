require './domaci'

parser = ExcelParser.new ("proba.xlsx")

puts parser.row(2)
puts parser.Prva_Kolona
puts parser.Prva_Kolona.b
