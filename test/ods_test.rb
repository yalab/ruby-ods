# -*- coding: utf-8 -*-
require 'test/unit'
require File.dirname(File.expand_path(__FILE__)) + '/../lib/ods'

class OdsTest < Test::Unit::TestCase
  BASE_DIR = File.dirname(File.expand_path(__FILE__))
  def setup
    @ods = Ods.new(BASE_DIR + '/cook.ods')
  end

  def teardown
    if File.exists?(BASE_DIR + '/modified.ods')
      File.unlink(BASE_DIR + '/modified.ods')
    end
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
    file_path = BASE_DIR + '/modified.ods'

    assert_not_equal modified_name, @ods.sheets[offset].name
    @ods.sheets[offset].name = modified_name
    @ods.save(file_path)
    modified_ods = Ods.new(file_path)
    assert_equal modified_name, modified_ods.sheets[offset].name
  end
end
