#encoding: utf-8
=begin rdoc
Make test example.
=end
WARNING = <<-txt
Warning! 

This file creator stores the data in a different way then excel.

The formula are not evaluated (they will get no value), so you have to 
open and save the file once with Excel.

txt

require 'rake'
require 'axlsx'
MyRand = Random.new(15) #Make randm number, but repeatable
VALUES = %w(first second third fourth fifth sixth)

file 'generated.xlsx' => [:warning, __FILE__] do |tsk|
  Axlsx::Package.new do |p| 
    wb = p.workbook
    wb.styles{|s|
      bold = s.add_style(:b => true)
      italic = s.add_style(:i => true)
      strike = s.add_style(:strike => true)
      bold_italic = s.add_style(:b => true, :i => true)
      time = s.add_style(:format_code => "YYYY-MM-DD hh:mm")
      wb.add_worksheet(:name => "DATA") do |sheet|
        sheet.add_row(
          ['String', 'Integer',  'Formula', 'Float', 'String with float value', 'Boolean (String > 0.5)', 'Date', 'Time'],
          :style => bold,
        )
        VALUES.each_with_index { |label,i| 
          sheet.add_row(
             [ #Values
                label,   
                MyRand.rand(24)+1, 
                "=B%i * 2" % [i+2], #Formula
                x = MyRand.rand, 
                x,    
                x > 0.5, 
                Date.new(2015,1,1) + i, 
                Time.new(2015,1,1,12) + i*60*60*24 + i*60*60,
              ],
            :types => [:string, :integer, nil, :float, :string, :boolean, :date, :time],
            :style  => [nil, nil, nil,nil,nil,nil, nil,time],
          )
        }
        #Only for printed sheets
        #~ wb.add_defined_name("'repeated header'!$1:$1", :local_sheet_id => sheet.index, :name => '_xlnm.Print_Titles')
      end
      wb.add_worksheet(:name => "Merged Areas") do |sheet|
        1.upto(15){|i|
          sheet.add_row( 'A'.upto('K').map{|char| '%s%i' % [char,i]})
        }
        #The data behind stay available.
        #so roo don't return nil for merged data.
        #-> The invisible columns must be set to nil.
        sheet.merge_cells("B2:G2")  #merged column -> \multicol
        sheet["C2:G2"].each{|cell| cell.value = nil } #Delete content of merged cells
        
        sheet.merge_cells("C4:C10") #merged row -> not supported
        sheet["C5:C10"].each{|cell| cell.value = nil } #Delete content of merged cells
        
        sheet.merge_cells("G8:I13")#Block -> not supported/\multicol
        sheet["H8:I13"].each{|cell| cell.value = nil } #Delete content of merged cells        
        sheet["G9:G13"].each{|cell| cell.value = nil } #Delete content of merged cells        
      end
      wb.add_worksheet(:name => "Formats") do |sheet|
        sheet.add_row(['Normal']) #Bold
        sheet.add_row(['Bold'],:style => bold) #Bold ##this bold is not set correct
        sheet.add_row(['Italic'],:style => italic) 
        sheet.add_row(['Strike'],:style => strike) 
        sheet.add_row(['Bold+Italic'],:style => bold_italic) 
        
        #This variants store boldin a different way:
        #~ sheet.add_row(['Bold'],:b => true) #Bold ##this bold is not set correct
        #~ sheet.add_row(['Italic'],:i => true) 
        #~ sheet.add_row(['Strike'],:strike => true) 
        #~ sheet.add_row(['Bold+Italic'],:b => true,:i => true) 
      end
    } #styles
    #~ wb.add_worksheet(:name => "References") do |sheet|
        #~ sheet.add_row ['Text', 'Reference', 'lookup-formula']
        #~ VALUES.each_with_index{ |label,i| 
          #~ #Correct example: =SVERWEIS(A1;DATA.A1:DATA.B3;2)
          #~ #There is a header line -> Offsets by 1
          #~ vlookup = '=VLOOKUP(A%i,DATA!A2:B%i,2)' % [i+2,VALUES.size+1]
          #~ sheet.add_row [label, vlookup, "'%s" % vlookup]
      #~ }
    #~ end
    puts "(Re)create %s" % tsk.name
    p.serialize(tsk.name)
  end
end

task :warning do
  puts WARNING
end

file 'simple.xlsx' => [__FILE__] do |tsk|
  Axlsx::Package.new do |p| 
    wb = p.workbook
    wb.add_worksheet(:name => "DATA") do |sheet|
      VALUES.each_with_index { |label,i| 
        sheet.add_row([label,   x = MyRand.rand(24)+1, x.odd?, MyRand.rand, Date.new(2015,1,1)],
            :types => [:string, :integer, :boolean, :float, :date]
        )
      }
    end
    wb.add_worksheet(:name => "DATA2") do |sheet|
      VALUES.reverse.each_with_index { |label,i| 
        sheet.add_row([label,   MyRand.rand(24)+1, MyRand.rand],
            :types => [:string, :integer, :float]
        )
      }
    end
    puts "(Re)create %s" % tsk.name
    p.serialize(tsk.name)
  end
end

#~ task :default => :warning
#~ task :default=> 'generated.xlsx'
task :default=> 'simple.xlsx'

Rake.application[:default].invoke if $0 == __FILE__

__END__
