= README

release::	0.0.0
copyright::	copyright(c) 2006-2009 ya-lab.org all rights reserved.


== About yalab-ruby-ods

This library is read and write OpenOffice Document SpreadSheet(ods) file.
This version only use string of cell format.


== Installation

    $ sudo gem install yalab-ruby-ods -s http://gemcutter.org

== Usage

    require 'rubygems'
    require 'ods'

    ods = Ods.new('some_document.ods')

    sheet = ods.sheets[0]
    sheet[3, :A].text #=> get A3 cell value
    sheet[4, :B].text = 'foobar'
    sheet[4, :B].text #=> foobar

    values = []
    sheet.rows.each do |row|
      row.each{|cell|
        values.push cell.text
      }
    end

    new_sheet = ods.create_sheet
    new_sheet[1, :A].annotation = 'hint'
    new_sheet[1, :A].text = 'baz'

    ods.save

== License

MIT License


== Author

yalab <rudeboyjet@gmail.com>
