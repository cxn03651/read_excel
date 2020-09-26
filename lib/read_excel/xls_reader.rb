# -*- conding: utf-8 -*-
require 'spreadsheet'

class ReadExcel::XLSReader
  class Worksheet
    def initialize(worksheet)
      @worksheet = worksheet
    end

    def name
      @worksheet.name
    end

    def row(n)
      values = []
      if @worksheet.row(n)
        count = @worksheet.row(n).size
        (0...count).each do |col|
          cell = @worksheet.row(n)[col]
          if cell.class == Spreadsheet::Formula
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
    @workbook = Spreadsheet.open(filepath)
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
