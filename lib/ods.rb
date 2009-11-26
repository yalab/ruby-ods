# -*- coding: utf-8 -*-
require 'rubygems'
require 'nokogiri'
require 'zip/zip'
require 'fileutils'

Nokogiri::XML::Element.module_eval do
  def add_element(name, attributes={})
    (prefix, name) = name.split(':') if name.include?(':')
    node = Nokogiri::XML::Node.new(name, self)
    attributes.each do |attr, val|
      node.set_attribute(attr, val)
    end
    ns = node.add_namespace_definition(prefix, Ods::NAMESPACES[prefix])
    node.namespace = ns
    self.add_child(node)
    node
  end

  def fetch(xpath)
    if node = self.xpath(xpath).first
      return node
    end

    return self.add_element(xpath) unless xpath.include?('/')

    xpath = xpath.split('/')
    last_path = xpath.pop
    fetch(xpath.join('/')).fetch(last_path)
  end
end

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
                      'number-columns-repeated' => '2',
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
      rows = @content.xpath('./table:table-row').to_a
      (row - rows.length).times do
        rows.push @content.add_element('table:table-row',
                                       'table:style-name' => 'ro1')
      end
      row = rows[row-1]
      col = ('A'..col.to_s).to_a.index(col.to_s)
      cols = row.xpath('table:table-cell').to_a
      (col - cols.length + 1).times do
        cols.push row.add_element('table:table-cell', 'office:value-type' => 'string')
      end
      Cell.new(cols[col])
    end
  end

  class Cell
    def initialize(content)
      @content = content
    end

    def text
      fetch('text:p').content
    end

    def text=(value)
      fetch('text:p').content = value
    end

    def annotation
      fetch('office:annotation/text:p').content
    end

    def annotation=(value)
      fetch('office:annotation/text:p').content = value
    end

    private
    def fetch(xpath)
      @content.fetch(xpath)
    end
  end
end
