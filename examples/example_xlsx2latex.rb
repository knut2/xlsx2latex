#encoding: utf-8
=begin rdoc
Example for usage of gem xlsx2latex
=end
$:.unshift('../lib')
require 'xlsx2latex'

require 'docgenerator'
require 'rake4latex'

=begin rdoc

Insipration:
http://softwarerecs.stackexchange.com/questions/2428/program-to-convert-excel-2013-spreadsheets-to-latex-tables

https://github.com/roo-rb/roo

Similar:
http://www.ctan.org/pkg/excel2latex
Excel-Makros to build LaTeX-table. 

excel2latex_Testdaten.xlsx
3 sheets
* 1: Numbers
* 2: Header line
* 3: multicol + multirow
=end

file 'test.tex' => __FILE__
file 'test.tex' => '../lib/xlsx2latex/excel.rb'
file 'test.tex' => "../unittest/testdata/testdata_excel2latex.xlsx" do |tsk| 
  doc = Docgenerator::Document.new(:template => :article_utf8)
  excel = XLSX2LaTeX::Excel.new("../unittest/testdata/testdata_excel2latex.xlsx", float_format: '%0.0f')
  puts excel.info
  doc.body << element(:section,{},'Sheet 1').Cr
  doc.body << excel.to_latex(sheetname: 'Sheet1')
  doc.body << element(:section,{},'Sheet2').Cr
  doc.body << excel.to_latex(sheetname: 'Sheet2')
  doc.save(tsk.name)
end

file 'test2.tex' => __FILE__
file 'test2.tex' => '../lib/xlsx2latex/excel.rb'
file 'test2.tex' => "../unittest/testdata/generated.xlsx" do |tsk|
  
  doc = Docgenerator::Document.new(:template => :article_utf8)
  doc.add_option('landscape')
  excel = XLSX2LaTeX::Excel.new("../unittest/testdata/generated.xlsx", float_format: '%0.0f')
  #~ excel.log.level = Log4r::DEBUG
  doc.body << element(:section,{},'Data overview').Cr
  doc.body << excel.to_latex(sheetname: 'DATA')
  doc.body << element(:section,{},'usage of formats').Cr
  doc.body << excel.to_latex(sheetname: 'Formats')
  doc.body << element(:section,{},'Merged Areas').Cr
  doc.body << excel.to_latex(sheetname: 'Merged Areas')
  doc.save(tsk.name)
end


#~ task :default => 'test.pdf'
task :default => 'test2.pdf'
task :default => :clean
Rake.application[:default].invoke if __FILE__== $0

