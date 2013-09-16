require 'test_helper'

class WorkbookTest < Test::Unit::TestCase
  include Sheet2hash::Errors
  
  def setup
    @header = Sheet2hash::Workbook.new File.expand_path('../header.xls', __FILE__)
    @no_header = Sheet2hash::Workbook.new File.expand_path('../no-header.xls', __FILE__)
    @dup_header = Sheet2hash::Workbook.new File.expand_path('../dup-header.xls', __FILE__)
  end
  
  def test_turn_to
    @header.turn_to 'Sheet1'
    assert_equal 'Sheet1', @header.sheet
    @header.turn_to 1
    assert_equal 'Sheet1', @header.sheet
  end
  
  def test_turn_to_error
    assert_raise SheetNotFoundError do
      @header.turn_to 'No Sheet'
    end
  end
  
  def test_sheets
    assert_equal ['Sheet1'], @header.sheets
  end
  
  def test_sheet
    assert_equal 'Sheet1', @header.sheet
  end
  
  def test_to_h
    assert_equal({ 'Sheet1' => [{ 'col1' => 1, 'col2' => 2, 'col3' => 3 }, { 'col1' => 'one', 'col2' => 'two', 'col3' => 'three' }] }, @header.to_h)
  end
  
  def test_to_a
    assert_equal [{ 'col1' => 1, 'col2' => 2, 'col3' => 3 }, { 'col1' => 'one', 'col2' => 'two', 'col3' => 'three' }], @header.to_a
  end
  
  def test_option_header1
    @header.turn_to 1, header: [:col1, :col2]
    assert_equal [{ col1: 1, col2: 2 }, { col1: 'one', col2: 'two' }], @header.to_a
  end
  
  def test_option_header2
    @no_header.turn_to 1, header: 2
    assert_equal [{ 'one' =>  1, 'two' => 2, 'three' => 3 }], @no_header.to_a
  end
  
  def test_option_start
    @header.turn_to 1, start: 3
    assert_equal [{ 'col1' => 'one', 'col2' => 'two', 'col3' => 'three' }], @header.to_a
    @header.turn_to 1, start: 4
    assert_equal [], @header.to_a
  end
  
  def test_option_end
    @header.turn_to 1, end: 2
    assert_equal [{ 'col1' => 1, 'col2' => 2, 'col3' => 3 }], @header.to_a
    @header.turn_to 1, end: 4
    assert_equal [{ 'col1' => 1, 'col2' => 2, 'col3' => 3 }, { 'col1' => 'one', 'col2' => 'two', 'col3' => 'three' }], @header.to_a
  end
  
  def test_option_keep_row
    @header.turn_to 1, keep_row: 1
    assert_equal [{ 'col1' => 'col1', 'col2' => 'col2', 'col3' => 'col3' }], @header.to_a
  end
  
  def test_option_skip_row
    @header.turn_to 1, skip_row: 3
    assert_equal [{ 'col1' => 1, 'col2' => 2, 'col3' => 3 }], @header.to_a
  end
  
  def test_option_keep_col
    @header.turn_to 1, keep_col: 1
    assert_equal [{ 'col1' => 1 }, { 'col1' => 'one' }], @header.to_a
  end
  
  def test_option_skip_col
    @header.turn_to 1, skip_col: [2, 3]
    assert_equal [{ 'col1' => 1 }, { 'col1' => 'one' }], @header.to_a
  end
  
  def test_uniq_unique_header
    assert_equal [{ 'col' => 1, 'col_dup' => 2, 'col_dup2' => 3 }, { 'col' => 'one', 'col_dup' => 'two', 'col_dup2' => 'three' }], @dup_header.to_a
  end
end
