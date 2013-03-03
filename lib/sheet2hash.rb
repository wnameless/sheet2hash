require 'sheet2hash/version'
require 'sheet2hash/railtie' if defined?(Rails)
require 'roo'

module Sheet2hash
  class Instance
    def initialize(sheet)
      @sheet = sheet
    end
    
    def turn_to(name)
      @sheet.default_sheet = name
      self
    end
    
    def sheets
      @sheet.sheets
    end
    
    def to_h
      ary_of_hash = sheet2ary_of_hash @sheet
      hash = {}
      ary_of_hash.each_with_index { |h, i| hash[i] = h }
      hash
    end
    
    private
    def sheet2ary_of_hash(s)
      ary_of_hash = []
      header = []
      (s.first_column..s.last_column).each { |col| header << s.cell(1, col) }
      (s.first_row..s.last_row).drop(1).each do |row|
        record = []
        (s.first_column..s.last_column).each { |col| record << s.cell(row, col) }
        ary_of_hash << Hash[ header.zip record ]
      end
      ary_of_hash
    end
  end
  
  def sheet2hash(data)
    path = save_sheet data
    sheet = Roo::Spreadsheet.open path
    File.delete path
    Instance.new sheet
  end
  
  private
  def save_sheet(data)
    name = data.original_filename
    directory = 'tmp'
    # create the file path
    path = File.join(directory, name)
    # write the file
    File.open(path, "wb") { |f| f.write(data.read) }
    path
  end
end
