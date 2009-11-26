# -*- coding: utf-8 -*-
require 'test/unit'
require File.dirname(File.expand_path(__FILE__)) + '/../lib/ods'

class OdsTest < Test::Unit::TestCase
  BASE_DIR = File.dirname(File.expand_path(__FILE__))
  def setup
    @ods = Ods.new(BASE_DIR + '/cook.ods')
    @file_path = BASE_DIR + '/modified.ods'
  end

  def teardown
    File.unlink(@file_path) if File.exists?(@file_path)
  end

  def test_sheet_count
    assert_equal 3, @ods.sheets.length
  end

  def test_sheet_name
    assert_equal 'さつま揚げとキャベツのお味噌汁', @ods.sheets[0].name
  end

  def test_sheet_name_modify
    modified_name = 'hogehoge'
    offset = 2

    assert_not_equal modified_name, @ods.sheets[offset].name
    @ods.sheets[offset].name = modified_name
    @ods.save(@file_path)
    modified_ods = Ods.new(@file_path)
    assert_equal modified_name, modified_ods.sheets[offset].name
  end

  def test_get_column
    sheet = @ods.sheets[0]
    assert_equal 'だし汁', sheet[2, :A].text
    assert_equal '適量', sheet[6, :B].text
  end

  def test_modify_column
    sheet_offset = 0
    row = 2
    col = :B
    sheet = @ods.sheets[sheet_offset]
    modified_text = '酢味噌'
    assert_not_equal modified_text, sheet[row, col].text
    sheet[row, col].text = modified_text
    @ods.save(@file_path)
    modified_ods = Ods.new(@file_path)
    assert_equal modified_text, modified_ods.sheets[sheet_offset][row, col].text
  end

  def test_access_not_existed_sheet
    ods_length = @ods.sheets.length
    new_sheet = @ods.create_sheet
    assert_equal "Sheet#{ods_length+1}", new_sheet.name
    assert_equal '', new_sheet[1, :A].text
    (col, row) = [100, :CC]
    assert_nothing_raised { new_sheet[col, row].text = 'hoge' }
    assert_equal 'hoge', new_sheet[col, row].text
  end

  def test_read_annotation
    cell = @ods.sheets[0][2, :A]
    assert_equal '昆布だし', cell.annotation
  end

  def test_write_annotation
    sheet_offset = 0
    row = 3
    col = :A
    cell = @ods.sheets[sheet_offset][row, col]
    assert_equal '', cell.annotation
    text = 'foobar'
    cell.annotation = text
    assert_equal text, @ods.sheets[sheet_offset][row, col].annotation
  end

  def test_columns_repeated
    sheet = @ods.create_sheet
    row = 10
    col = :C
    sheet[row, col].text = 'hoge'
    @ods.save(@file_path)

    modified_ods = Ods.new(@file_path)
    sheet = modified_ods.sheets[modified_ods.sheets.length-1]
    assert_equal "3", sheet.column.attr('repeated')
  end

  def test_each_rows
    sheet = @ods.create_sheet
    rows = 10
    cols = :A
    sheet[rows, cols].text = 'foo'

    count = 0
    sheet.rows.each do |row|
      count += 1
    end
    assert_equal rows, count
  end
end
