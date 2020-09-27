# ReadExcel

This ReadExcel rubygem can handle existing excel file both *.xlsx and *.xls. It can only get cell's value and sheet name of excel file.

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'read_excel'
```

And then execute:

    $ bundle install

Or install it yourself as:

    $ gem install read_excel

## Usage

```ruby
require 'read_excel'

workbook = ReadExcel.new('/path/to/excel_file')

sheet_count = workbook.worksheets.count

worksheets = workbook.worksheets

worksheets.each do |worksheet|
  p worksheet.name
end

worksheet = worksheets[0]

row_values = worksheet.row(0)  # (row, col) = (0, 0) = 'A1'
p row_values[0]  # 'A1'
p row_values[2]  # 'C1'

row_values = worksheet.row(10000)
row_values.empty? # => true   empty row returns []
```

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/[USERNAME]/read_excel. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [code of conduct](https://github.com/[USERNAME]/read_excel/blob/master/CODE_OF_CONDUCT.md).


## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the ReadExcel project's codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](https://github.com/[USERNAME]/read_excel/blob/master/CODE_OF_CONDUCT.md).
