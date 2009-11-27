# -*- coding: utf-8 -*-
require 'forwardable'
require 'rubygems'
require 'nokogiri'
require 'nokogiri_ext'
require 'zip/zip'
require 'fileutils'

class Ods
  attr_reader :content, :sheets
  XPATH_SHEETS = '//office:body/office:spreadsheet/table:table'

  NAMESPACES = {
    'office' => 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    'style'  => 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
    'text'   => 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'table'  => 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    'draw'   => 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
    'fo'     => 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
    'xlink'  => 'http://www.w3.org/1999/xlink',
    'dc'     => 'http://purl.org/dc/elements/1.1/',
    'meta'   => 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
    'number' => 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
    'presentation' => 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0',
    'svg'    => 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
    'chart'  => 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0',
    'dr3d'   => 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0',
    'math'   => 'http://www.w3.org/1998/Math/MathML',
    'form'   => 'urn:oasis:names:tc:opendocument:xmlns:form:1.0',
    'script' => 'urn:oasis:names:tc:opendocument:xmlns:script:1.0',
    'ooo'    => 'http://openoffice.org/2004/office',
    'ooow'   => 'http://openoffice.org/2004/writer',
    'oooc'   => 'http://openoffice.org/2004/calc',
    'dom'    => 'http://www.w3.org/2001/xml-events',
    'xforms' => 'http://www.w3.org/2002/xforms',
    'xsd'    => 'http://www.w3.org/2001/XMLSchema',
    'xsi'    => 'http://www.w3.org/2001/XMLSchema-instance',
    'rpt'    => 'http://openoffice.org/2005/report',
    'of'     => 'urn:oasis:names:tc:opendocument:xmlns:of:1.2',
    'rdfa'   => 'http://docs.oasis-open.org/opendocument/meta/rdfa#',
    'field'  => 'urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0',
    'formx'  => 'urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0'
  }

  def initialize(path)
    @path = path
    Zip::ZipFile.open(@path) do |zip|
      @content = Nokogiri::XML::Document.parse(zip.read('content.xml'))
    end
    @sheets = []
    @content.root.xpath(XPATH_SHEETS).each do |sheet|
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

    @sheets.each do |sheet|
      column = sheet.column
      max_length = 0
      column.content.parent.xpath('table:table-row').each do |row|
        length = row.xpath('table:table-cell').length
        max_length = length if max_length < length
      end
      column.set_attr('repeated', max_length)
    end

    Zip::ZipFile.open(dest) do |zip|
      zip.get_output_stream('content.xml') do |io|
        io << @content.to_s
      end
    end
  end

  def create_sheet
    parent = @content.root.xpath(XPATH_SHEETS.split('/')[0..-2].join('/'))[0]
    table = parent.add_element('table:table',
                               'name'       => "Sheet#{@sheets.length + 1}",
                               'style-name' => 'ta1',
                               'print'      => 'false')
    table.add_element('table:table-column',
                      'style-name'              => 'co1',
                      'default-cell-style-name' => 'Default')
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
      @content.set_attribute('table:name', name)
    end

    def [](row, col)
      (row - rows.length).times do
        rows.push(Row.new(@content.add_element('table:table-row',
                                               'table:style-name' => 'ro1'), rows.length+1))
      end
      row = rows[row-1]
      col = ('A'..col.to_s).to_a.index(col.to_s)
      cols = row.cols
      (col - cols.length + 1).times do
        cols.push row.create_cell
      end
      cols[col]
    end

    def rows
      return @rows if @rows
      @rows = []
      @content.xpath('./table:table-row').each_with_index{|row, index|
        @rows << Row.new(row, index+1)
      }
      @rows
    end

    def column
      Column.new(@content.xpath('table:table-column').first)
    end
  end

  class Row
    extend Forwardable

    def_delegator :@content, :xpath, :xpath
    def_delegator :@content, :add_element, :add_element
    attr_reader :no

    def initialize(content, no)
      @content = content
      @no = no
    end

    def cols
      @cols ||= xpath('table:table-cell').map{|cell| Cell.new(cell)}
    end

    def create_cell
      Cell.new(@content.add_element('table:table-cell', 'office:value-type' => 'string'))
    end
  end

  class Cell
    extend Forwardable

    def_delegator :@content, :fetch, :fetch

    def initialize(content)
      @content = content
    end

    def value
      fetch('text:p').content
    end

    def value=(value)
      fetch('text:p').content = value
    end

    def annotation
      fetch('office:annotation/text:p').content
    end

    def annotation=(value)
      fetch('office:annotation/text:p').content = value
    end
  end

  class Column
    attr_reader :content
    def initialize(content)
      @content = content
    end

    def attr(name)
      @content['number-columns-' + name]
    end

    def set_attr(name, value)
      @content['table:number-columns-' + name] = value.to_s
    end
  end
end
