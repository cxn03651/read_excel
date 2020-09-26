# coding: utf-8
require "test_helper"

class ReadExcelXLSReaderTest < Minitest::Test
  def setup
    @filepath = File.join(File.dirname(File.expand_path(__FILE__)), 'testfile.xls')
    @workbook = nil
  end

  def test_worksheets
    assert workbook.worksheets
  end

  def test_worksheet_count
    assert_equal(2, workbook.worksheets.count)
  end

  def test_worksheet_index
    count = workbook.worksheets.count
    (0...count).each do |i|
      assert workbook.worksheets[i]
    end
  end

  def test_worksheet_name
    assert_equal('Sheet1', workbook.worksheets[0].name)
    assert_equal('シート２', workbook.worksheets[1].name)
  end

  # n行目のvalueのArrayを返す
  def test_worksheet_row
    worksheet = workbook.worksheets[0]

    values = worksheet.row(0)
    assert_equal(1, values.count)
    assert_equal(
      ['ReadExcel gemのテスト用'],
      values
    )

    values = worksheet.row(1)
    assert_equal(8, values.count)
    assert_equal(
      ['a', '日本語', 1, 0.5, 0.5, nil],
      values[0..5]
    )

    values = worksheet.row(2)
    assert_equal(0, values.count)
    assert_equal(
      [],
      values
    )

    values = worksheet.row(3)
    assert_equal(8, values.count)
    assert_equal(
      ['a', '日本語', 1, 0.5, 0.5, nil],
      values[0..5]
    )

    values = worksheet.row(4)
    assert_equal(0, values.count)
    assert_equal(
      [],
      values
    )
  end

  private

    def workbook
      @workbook ||= ReadExcel::XLSReader.new(@filepath)
    end
end
