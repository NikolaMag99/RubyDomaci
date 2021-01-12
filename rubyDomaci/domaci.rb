require 'roo'

class Column

    def initialize(column, colindex, table)
        @column = column
        @colindex = colindex
        @table = table
        
    end

    def sum
        @column.inject(0){ |sum, x| sum + x.to_f }
    end

    def method_missing(m)
        @column.each_with_index do |col, i|
            if col.to_s.downcase.gsub(' ', '-').eql?(m.to_s.downcase.gsub(' ', '-')) then
                return @table.row(i+2)
            end
        end
    end

end


class Parser
    include Enumerable

    def initialize(path)
        @table = Roo::Spreadsheet.open(path, {:expand_merged_ranges => true})
        @columns = Hash.new
        @table.row(1).each_with_index { |header, i| @columns[header.downcase.gsub(' ', '-')] = i+1 }
    end

    def method_missing(header)
        Column.new(@table.column(@columns[header.to_s]).drop(1), @table, @columns[header.to_s])
    end

    def [](key)
        return @table.column(@columns[key.to_s.downcase.gsub(' ','-')]).drop(1)
    end

    def row(num)
        return @table.row(num)
    end

    def each
        yield @table.parse
    end

    def to_s
        @table.parse
    end
end