#!/usr/bin/env ruby
#encoding: utf-8
=begin rdoc
Gem XLSX2LaTeX: Convert excel-files to LaTeX-tabulars.
=end
require_relative '../lib/xlsx2latex'

require 'optparse'

options = {}

opt_parser = OptionParser.new do |opt|
  opt.banner = "Usage: %s excelfile [OPTIONS]" % __FILE__
  opt.separator  ""
  opt.separator  "Options"

  opt.on("-t","--target TARGET","target filename (default STDOUT)") do |target|
    options[:target] = target
  end

  opt.on("-s","--sheet SHEET","Source sheet in Excel (default: all)") do |sheet|
    options[:sheet] = sheet
  end

  opt.on("-f","--float FLOAT","Format for floats") do |float|
    options[:float] = float
  end
  
  opt.on("-d","--date DATE","Format for dates") do |date|
    options[:date] = date
  end
  
  opt.on("--time TIME","Format for times") do |time|
    options[:time] = time
  end
  
  opt.on("--true TRUE","Format for true") do |vtrue|
    options[:true] = vtrue
  end
  opt.on("--false FALSE", "Format for false") do |vfalse|
    options[:false] = vfalse
  end
  
=begin  
Generate missing options

no_format no_bold no_italics...
bool: text, surd...
all_sheets
       
=end
  %w{}.each{|missing|
  puts <<code % [missing[0], missing, missing.upcase, missing, missing, missing, missing]
  opt.on("-%s","--%s %s","%s") do |%s|
    options[:%s] = %s
  end
code
}

  opt.on("-h","--help","help") do
    puts opt_parser
    puts <<examples

==Examples:
Write default data sheet as excel to stdout:
  xlsx2latex my_excel.xlsx
  
Write sheet with name "DATA2" as excel to stdout:
  xlsx2latex my_excel.xlsx -s DATA2

No decimals for all floats (numbers):
  xlsx2latex my_excel.xlsx -f %.0f

Change data format (e.g. to "1. January 2015"):
  xlsx2latex my_excel.xlsx -d "%d. %B %Y"

Change format for boolean values:
  xlsx2latex my_excel.xlsx  --true yes --false no

examples
    exit
  end
end

#Quick test
if $0 ==   __FILE__  and File.basename($0) == $0
  ARGV << '../unittest/testdata/simple.xlsx'  #source
  #~ ARGV << '-t' << 'test.tex'
  #~ ARGV << '-s' << 'DATA2'
  #~ ARGV << '-f' << '%09.3f'
  #~ ARGV << '--true' << 'yes'
  #~ ARGV << '--false' << 'no'
  #~ ARGV << '-d' << '%Y'
  ARGV << '-h' 
end

opt_parser.parse!
 if ! ARGV[0]
  puts "No Excel file given. Try %s -h for more information" % __FILE__
  exit 1 
end
if ! File.exist?(ARGV[0])
  puts "Excel file %s not found" % ARGV[0]
  exit 2 
end

$excel = XLSX2LaTeX::Excel.new(ARGV[0])
texcode = <<preamble % [File.basename(__FILE__), ARGV[0] ]
%%encoding: utf-8
%%%%%%%%%%%%%%%%%%
%%
%% LaTeX tabular created by %s
%% Source: %s
%%
%%%%%%%%%%%%%%%%%%
preamble
sheets = []
if options[:sheet]
  sheets << options[:sheet] 
else
  #~ sheets << $excel.xlsx.default_sheet
  sheets = $excel.xlsx.sheets
end

sheets.each do |sheetname|
  texcode << "\n\n%%\n%%Source: %s [%s]\n%%\n" % [ ARGV[0], sheetname]
  texcode << $excel.to_latex(
    sheetname: sheetname,
    float_format: options[:float] || '%0.2f',
    date_format: options[:date] || '%Y-%m-%d',
    time_format:  options[:time] || '%Y-%m-%d %H:%M',
    bool_true: options[:true] || 'X',
    bool_false: options[:false] || '-',
  )
end

if options[:target]
  File.open(options[:target], 'w:utf-8'){|f| 
    f << texcode
    $excel.log.info("Created %s" % options[:target])
  }
else
  puts texcode
end

__END__
