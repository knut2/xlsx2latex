#encoding: utf-8
=begin
Create gem xlsx2latex
=end
$:.unshift('c:/usr/script/knut/knut-testtask/lib')
$:.unshift('c:/usr/script/knut/knut-gempackager/lib')
require 'knut-gempackager'  

$:.unshift('lib')
require 'xlsx2latex'
$xlsx2latex_version = "0.1.0.rc3"

#http://docs.rubygems.org/read/chapter/20
gem_xlsx2latex = Knut::Gem_packer.new('xlsx2latex', $xlsx2latex_version){ |gemdef, s|
  s.name = "xlsx2latex"
  s.version =  $xlsx2latex_version
  s.author = "Knut Lickert"
  s.email = "knut@lickert.net"
  s.homepage = "https://github.com/knut2/xlsx2latex"
  #~ s.homepage = "https://rubygems.org/gems/xlsx2latex"
  s.platform = Gem::Platform::RUBY
  s.required_ruby_version = '>= 2.0' #uses new interface
  s.license = 'LPPL-1.3c'
  s.summary = "Convert Excel-file to LaTeX-tabulars"
  s.description = <<-DESCR
Take an Excel-file (xlsx) and convert the content into a LaTeX-tabular.

By default each tab becomes a tabular.
  DESCR
  s.require_path = "lib"
  s.files = %w{
    readme.rdoc
    bin/xlsx2latex.rb
    lib/xlsx2latex.rb
    lib/xlsx2latex/excel.rb
    examples/example_xlsx2latex.rb
  }
  s.test_files    = %w{
    unittest/test_testdata_excel2latex.rb
    unittest/test_generated_xlsx.rb
    unittest/test_bin.rb
    unittest/testdata/make_testdata.rb
    unittest/testdata/generated.xlsx
    unittest/testdata/simple.xlsx
    unittest/testdata/testdata_excel2latex.xlsx
  }
  s.test_files   << Dir['unittest/expected/*']
  s.test_files.flatten!

  s.bindir = "bin"
  s.executables << 'xlsx2latex.rb'

  s.rdoc_options << '--main' << 'readme.rdoc'
  s.extra_rdoc_files = %w{
    readme.rdoc
    bin/readme.txt
  }
  
  s.add_dependency('roo', '~> 2') 
  s.add_dependency('log4r', '~> 1')
  #~ s.add_dependency('optparse', '~> 0') #is a std.lib
  
  s.add_development_dependency('minitest-filecontent', '~> 0')#unit tests
  s.add_development_dependency('axlsx', '~> 2')         #Make test example
  #~ s.requirements << ''

  gemdef.public = true

  gemdef.define_test( 'unittest', FileList['test*.rb'])
  gemdef.versions << XLSX2LaTeX::VERSION

}

#generate rdoc
task :rdoc_local do
  FileUtils.rm_r('doc') if File.exist?('doc')
  cmd = ["rdoc -f hanna"]
  cmd << gem_xlsx2latex.spec.lib_files
  cmd << gem_xlsx2latex.spec.extra_rdoc_files
  cmd << gem_xlsx2latex.spec.rdoc_options
  `#{cmd.join(' ')}`
end
file 'bin/xlsx2latex.exe' => 'bin/xlsx2latex.rb' do
  Dir.chdir('bin'){
    puts "Create 'xlsx2latex.exe'"; STDOUT.flush
    `ocra xlsx2latex.rb`
  }
end
file 'bin/readme.txt' => __FILE__
file 'bin/readme.txt' => 'bin/xlsx2latex.rb' do |tsk|
  puts "Create 'xlsx2latex.txt'"; STDOUT.flush
  File.open(tsk.name, 'w:utf-8'){|f|
    f << <<readme
xlsx2LaTeX
===========
Convert Excel-sheets into LaTeX-tables.

You can take the LaTeX-table and continue with your work.

Some features:
* Formats for numbers, booleans, dates and times can be defined.
* Selection of datasheet is possible.

readme
    Dir.chdir('bin'){ f << `ruby xlsx2latex.rb -h`.gsub(/\.rb/, '.exe')}
    f << <<readme
=======
Not supported features:
* xls-files [1]
* deep control on text formats [2]
* Suppress text formats (make no bold...) [3]
* select areas or specific columns for export [3]
* column specific formats [3]
* Detect number format of Excel and use it for the output [4]

remarks:
[1] maybe possible in future - not tested yet.
[2] Not planned in future
[3] Not planned yet, but on the todo list.
[4] Not realized, but wanted

readme
}
end

#~ desc "Gem xlsx2latex"
#~ task :default => :check
#~ task :default => 'bin/xlsx2latex.exe'
#~ task :default => 'bin/readme.txt'
#~ task :default => :test
#~ task :default => :gem
#~ task :default => :install
#~ task :default => :rdoc_local
#~ task :default => :links
#~ task :default => :ftp_rdoc
#~ task :default => :push


if $0 == __FILE__
  app = Rake.application
  app[:default].invoke
end
__END__

Todos:
* Use of booktabs?