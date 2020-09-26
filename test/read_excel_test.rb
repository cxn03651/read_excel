# coding: utf-8
require "test_helper"

class ReadExcelTest < Minitest::Test
  def test_that_it_has_a_version_number
    refute_nil ::ReadExcel::VERSION
  end

  # エクセル以外のファイルは例外を発生させる
  def test_non_excel_file_raise_exception
    filename = 'example.oss'
    e = assert_raises RuntimeError do
      ReadExcel.new(filename)
    end
    assert_equal "#{filename} doesn't seem excel file.", e.message
  end

  def test_both_type_succeed
    ['testfile.xls', 'testfile.xlsx'].each do |file|
      filepath = File.join(File.dirname(File.expand_path(__FILE__)), file)
      workbook = ReadExcel.new(filepath)
      assert workbook.worksheets
      count = workbook.worksheets.count
      assert_equal(2, count)
      (0...count).each do |i|
        assert workbook.worksheets[i]
      end
      assert_equal('Sheet1',   workbook.worksheets[0].name)
      assert_equal('シート２', workbook.worksheets[1].name)

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
  end

end
