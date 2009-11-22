# -*- coding: utf-8 -*-
require 'test/unit'
require File.dirname(File.expand_path(__FILE__)) + '/../lib/ods'

class OdsTest < Test::Unit::TestCase
  def setup
    @ods = Ods.new(File.dirname(File.expand_path(__FILE__)) + '/cook.ods')
  end

  def test_sheet_count
    assert_equal 3, @ods.sheets.length
  end

  def test_sheet_name
    assert_equal 'さつま揚げとキャベツのお味噌汁', @ods.sheets[0].name
  end
end
