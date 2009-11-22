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

  class Sheet
    def initialize(content)
      @content = content
    end

    def name
      @content.attribute('name').to_s
    end

    def name=(name)
      @content.add_attribute('table:name', name)
    end
  end
end
