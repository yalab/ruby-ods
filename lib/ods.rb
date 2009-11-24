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
                               'table:style-name' => "ta1",
                               'table:print'      => "false")
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
      row = @content.get_elements('table:table-row')[row-1]
      return nil unless row
      column = row.get_elements('table:table-cell')[('A'..col.to_s).to_a.index(col.to_s)]
      column.get_elements('text:p').first.get_text
    end

    def [](row, col)
      text_node(row, col).to_s
    end

    def []=(row, col, value)
      return unless node = text_node(row, col)
      node.value = value
    end
  end
end
