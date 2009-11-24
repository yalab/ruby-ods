# -*- coding: utf-8 -*-
require 'rexml/document'
require 'rubygems'
require 'zip/zip'
require 'fileutils'

class Ods
  attr_reader :content, :sheets
  XPATH_SHEETS = 'office:body/office:spreadsheet/table:table'

  def initialize(path)
    @path = path
    Zip::ZipFile.open(@path) do |zip|
      @content = REXML::Document.new zip.read('content.xml')
    end
    @sheets = []
    @content.root.get_elements(XPATH_SHEETS).each do |sheet|
      @sheets.push(Sheet.new(sheet))
    end
    @content
  end

  def save(dest=nil)
    if dest
      FileUtils.cp(@path, dest)
    else
      dest = @path
    end

    Zip::ZipFile.open(dest) do |zip|
      zip.get_output_stream('content.xml') do |io|
        @content.write(io)
      end
    end
  end

  def create_sheet
    parent = @content.root.get_elements(XPATH_SHEETS.split('/')[0..-2].join('/'))[0]
    table = parent.add_element('table:table',
                               'table:name'       => "Sheet#{@sheets.length + 1}",
                               'table:style-name' => 'ta1',
                               'table:print'      => 'false')
    table.add_element('table:table-column',
                      'table:style-name'              => 'co1',
                      'table:number-columns-repeated' => '2',
                      'table:default-cell-style-name' => 'Default')
    new_sheet = Sheet.new(table)
    @sheets.push(new_sheet)
    new_sheet
  end

  class Sheet
    attr_reader :content
    def initialize(content)
      @content = content
    end

    def name
      @content.attribute('name').to_s
    end

    def name=(name)
      @content.add_attribute('table:name', name)
    end

    def text_node(row, col)
      rows = @content.get_elements('table:table-row')
      (row - rows.length).times do
        rows.push @content.add_element('table:table-row',
                                       'table:style-name' => 'ro1')
      end
      row = rows[row-1]

      col = ('A'..col.to_s).to_a.index(col.to_s)
      cols = row.get_elements('table:table-cell')
      (col - cols.length + 1).times do
        cols.push row.add_element('table:table-cell', 'office:value-type' => 'string')
      end
      column = cols[col]

      unless cell = column.get_elements('text:p').first
        cell = column.add_element('text:p')
        cell.add_text('')
      end
      cell.get_text
    end

    def [](row, col)
      text_node(row, col).to_s
    end

    def []=(row, col, value)
      text_node(row, col).value = value
    end
  end
end
