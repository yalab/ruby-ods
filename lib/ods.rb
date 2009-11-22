require 'rexml/document'
require 'rubygems'
require 'zip/zip'


class Ods
  attr_reader :content, :sheets

  def initialize(path)
    Zip::ZipFile.open(path) do |zip|
      @content = REXML::Document.new zip.read('content.xml')
    end
    @sheets = []
    @content.root.get_elements('office:body/office:spreadsheet/table:table').each do |sheet|
      @sheets.push(Sheet.new(sheet))
    end
  end

  class Sheet
    def initialize(content)
      @content = content
    end

    def name
      @content.attribute('name').to_s
    end
  end
end
