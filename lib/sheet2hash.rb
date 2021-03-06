require 'roo'
require 'sheet2hash/errors'
require 'sheet2hash/options'

# Sheet2hash
# @author Wei-Ming Wu
module Sheet2hash
  # Sheet2hash::Workbook converts Excel or Spreadsheet into Ruby Hash.
  # @author Wei-Ming Wu
  class Workbook
    include Options, Errors
    
    # Creates a Workbook.
    #
    # @return [Workbook] a Workbook object
    def initialize path, opts = {}
      @workbook = Roo::Spreadsheet.open path
      @opts = process_options opts
      @sheet_opts = {}
      set_sheet_attributes
    end
    
    # Turns to specified sheet of this Workbook.
    #
    # @param sheet [String, Integer] the name or index(1-based) of a sheet
    # @param sheet_opts [Hash] the options of turn_to
    # @option sheet_opts [Array, Integer] :header using an Array to store the header of this sheet or choosing an index(1-based) of rows as the header
    # @option sheet_opts [Integer] :start the index(1-based) of rows to start
    # @option sheet_opts [Integer] :end the index(1-based) of rows to end
    # @option sheet_opts [Integer] :keep_row the indices(1-based) of rows to keep
    # @option sheet_opts [Integer] :skip_row the indices(1-based) of rows to skip
    # @option sheet_opts [Integer] :keep_col the indices(1-based) of columns to keep
    # @option sheet_opts [Integer] :skip_col the indices(1-based) of columns to skip
    def turn_to sheet, sheet_opts = {}
      if sheet.kind_of?(Integer) && sheet <= sheets.size
        @workbook.default_sheet = sheets[sheet - 1]
        @sheet_opts = process_options sheet_opts
        set_sheet_attributes
      elsif sheet.kind_of?(String) && sheets.include?(sheet)
        @workbook.default_sheet = sheet
        @sheet_opts = process_options sheet_opts
        set_sheet_attributes
      else
        raise SheetNotFoundError
      end
    end
    
    # Returns all sheets of this Workbook.
    #
    # @return [Array] all sheets of this Workbook
    def sheets
      @workbook.sheets
    end
    
    # Returns the name of current sheet.
    #
    # @return [String]  the name of current sheet
    def sheet
      @workbook.default_sheet
    end
    
    # Converts all sheets to a Hash, sheet names are the keys, values are Array of Hash.
    #
    # @return [Hash] sheet names are the keys, values are Array of Hash
    def to_h
      hash = {}
      sheets.each do |sheet|
        turn_to sheet
        hash[sheet] = to_a
      end
      hash
    end
    
    # Converts current sheet to an Array of Hash.
    #
    # @param [Array] Array of Hash
    def to_a
      ary = []
      @rows.each do |row|
        record = []
        @columns.each { |col| record << trim_int_cell(@workbook.cell(row, col)) }
        ary << Hash[ @header.zip record ]
      end
      ary
    end
    
    private
    
    def set_sheet_attributes
      @header = header
      @rows = rows
      @columns = columns
    end
    
    def trim_int_cell cell
      cell.kind_of?(Numeric) && cell % 1 == 0 ? cell.to_i : cell
    end
    
    def columns
      columns = (@workbook.first_column..@workbook.last_column).to_a
      if @sheet_opts[:skip_col] || @sheet_opts[:keep_col]
        collect_columns columns, @sheet_opts
      elsif @opts[:skip_col] || @opts[:keep_col]
        collect_columns columns, @opts
      else
        columns
      end
    end
    
    def collect_columns columns, opts
      columns.keep_if { |col| opts[:keep_col].include? col } if opts[:keep_col]
      columns = columns - opts[:skip_col] if opts[:skip_col]
      columns
    end
    
    def first_row
      first_row = @sheet_opts[:start] || @opts[:start]
      first_row || @workbook.first_row
    end
    
    def last_row
      last_row = @sheet_opts[:end] || @opts[:end] || @workbook.last_row
      last_row > @workbook.last_row ? @workbook.last_row : last_row
    end
    
    def rows
      rows = (first_row..last_row).to_a
      if @sheet_opts[:keep_row] || @sheet_opts[:skip_row]
        rows = collect_rows rows, @sheet_opts
        rows - [@sheet_opts[:header]]
      elsif @opts[:keep_row] || @opts[:skip_row]
        rows = collect_rows rows, @opts
        rows - [@opts[:header]]
      elsif @sheet_opts[:header].kind_of?(Integer) || @sheet_opts[:header].kind_of?(Integer)
        rows - [@sheet_opts[:header] || @sheet_opts[:header]]
      else
        rows - [@workbook.first_row]
      end
    end
    
    def collect_rows rows, opts
      rows.keep_if { |row| opts[:keep_row].include? row } if opts[:keep_row]
      rows = rows - opts[:skip_row] if opts[:skip_row]
      rows = rows - [@workbook.first_row] unless opts[:keep_row] && opts[:keep_row].include?(1)
      unless opts[:header].kind_of?(Integer) && opts[:keep_row].include?(opts[:header])
        rows = rows - [opts[:header]]
      end
      rows
    end
    
    def header
      header = @sheet_opts[:header] || @opts[:header]
      if header
        header = process_header header
      else
        header = header_from_row @workbook.first_row
      end
      match_header_with_columns header
    end
    
    def match_header_with_columns header
      header.each_with_index.select { |h, i| columns.include?(i + 1) }.map { |i| i.first }
    end
    
    def process_header header
      if header.kind_of? Array
        unique_header header
      elsif header.kind_of? Integer
        header_from_row header
      else
        raise InvalidHeaderError
      end
    end
    
    def header_from_row row
      header = []
      (@workbook.first_column..@workbook.last_column).each { |col| header << @workbook.cell(row, col) }
      unique_header header
    end
    
    def unique_header header
      header.reverse!
      dup_header = header.dup
      header = header.map do |field|
        count = dup_header.count field
        if count == 2
          dup_header.delete_at dup_header.find_index(field)
          "#{field}_dup"
        elsif count > 1
          dup_header.delete_at dup_header.find_index(field)
          "#{field}_dup#{count - 1}"
        else
          field
        end
      end
      header.reverse!
    end
  end
end
