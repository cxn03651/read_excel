# -*- conding: utf-8 -*-
require 'rubyXL'

class ReadExcel::XLSXReader
  class Worksheet
    def initialize(worksheet)
      @worksheet = worksheet
    end

    def name
      @worksheet.sheet_name
    end

    def row(n)
      values = []
      if @worksheet[n]
        count = @worksheet[n].size
        (0...count).each do |col|
          cell = @worksheet[n][col]
          if cell
            values << cell.value
          else
            values << cell
          end
        end
      end

      values
    end
  end

  def initialize(filepath)
    @workbook = RubyXL::Parser.parse(filepath)
    @worksheets = []
  end

  def worksheets
    if @worksheets.empty?
      @workbook.worksheets.each do |worksheet|
        @worksheets << Worksheet.new(worksheet)
      end
    end

    @worksheets
  end

end
