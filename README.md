# Sheet2hash

Convert Excel or Spreadsheet to Ruby Hash

## Installation

Add this line to your application's Gemfile:

    gem 'sheet2hash'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install sheet2hash

## Usage

Basic:

    wb = Sheet2hash.workbook.new 'path/to/your/excel_or_spreadsheet'
    wb.to_a # convert default sheet to array of hash
            # [{ field1 => data, field2 => data, ...}, { field1 => data, field2 => data, ...}, ...]
            # each hash is a row of sheet
            # the first row is used to be the keys of hash
    wb.to_h # convert all sheets to hash
            # the output will be { 'name_of_sheet' => [{ field => data, ... }, ...] }
            
Advanced:

    # All sheets are count from 1, NOT 0
    wb.sheets                           # list all sheets in the workbook
    wb.sheet                            # show the name of current workbook
    wb.turn_to(2)                       # make sheet 2 as current sheet
    wb.turn_to('name_of_the_3rd_sheet') # make sheet 3 as current sheet

Options:

     # All rows, columns are count from 1, NOT 0
    wb = Sheet2hash.workbook.new 'path/to/your/excel_or_spreadsheet',
      :header   => 3,                     # use specific row as header 
      :header   => ['col1', 'col2', ...], # use self-defined header
      :start    => 4,                     # start at row 4
      :end      => 20,                    # end at row 20
      :keep_row => [4, 5],                # only convert row 4 and 5
      :skip_row => [6, 7],                # skip row 6 and 7
      :keep_col => [1],                   # only convert column 1
      :skip_col => [2, 3]                 # skip column 2 and 3
      
    wb.turn_to(2, :herder => 3, ...)      # global options can be overrided by passing options to specfic sheet
                                          # global options will still affect after turning to another sheet
      

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
