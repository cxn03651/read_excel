require "read_excel/version"
require "read_excel/xls_reader"
require "read_excel/xlsx_reader"

class ReadExcel
  class Error < StandardError; end

  def initialize(filepath)
    @filepath = filepath
    if @filepath =~ /xls\Z/i
      @read_adapter = ReadExcel::XLSReader.new(filepath)
    elsif @filepath =~ /xlsx\Z/i
      @read_adapter = ReadExcel::XLSXReader.new(filepath)
    else
      raise "#{filepath} doesn't seem excel file."
    end
  end

  # list of worksheets
  def worksheets
    @read_adapter.worksheets
  end

end
