#encoding: utf-8
=begin rdoc
Tests for xlsx2latex-binary.
=end
gem 'minitest'
require 'minitest/autorun'
require 'minitest/filecontent' 
  
$:.unshift('../lib')
require 'xlsx2latex'

class Test_bin < MiniTest::Test
  def setup
    @cmd = 'ruby ../bin/xlsx2latex.rb'
    @src = 'testdata/simple.xlsx'
    @cmd_src = [@cmd, @src].join(' ')
  end
  def test_bin_no_parameter
    `#{@cmd}`    
    assert_equal(1,$?.exitstatus)
  end
  def test_bin_unkown_source
    `#{@cmd} notexisting.xlsx`
    assert_equal(2,$?.exitstatus)
  end
  #Result is posted on screen
  #The output contains also the Log-Messages.
  def test_bin_default
    output = `#{@cmd_src}`
    assert_equal(0,$?.exitstatus)
    assert_equal_filecontent('expected/%s.tex' % __method__, output)
  end
  
  #Save result immediate into file
  def test_bin_with_target
    filename = 'temporary_testfile_please_delete_%s_%s.tex' % [__method__.hash, __method__]
    refute(File.exist?(filename), 'Testfile %s already exists' % filename)
    output = `#{@cmd_src} -t #{filename}`
    assert_equal(0,$?.exitstatus)
    assert(File.exist?(filename), "Target file is not created")
    assert_equal(<<LOG % filename, output)
 INFO simple.xlsx: Convert Sheet DATA to LaTeX
 INFO simple.xlsx: Convert Sheet DATA2 to LaTeX
 INFO simple.xlsx: Created %s
LOG
    assert_equal_filecontent('expected/%s.tex' % __method__, File.read(filename))
    File.delete(filename)
  end

  def test_bin_select_sheet1
    output = `#{@cmd_src} -s DATA`
    assert_equal(0,$?.exitstatus)
    assert_equal_filecontent('expected/%s.tex' % __method__, output)
  end
  def test_bin_select_sheet2
    output = `#{@cmd_src} -s DATA2`
    assert_equal(0,$?.exitstatus)
    assert_equal_filecontent('expected/%s.tex' % __method__, output)
  end

  def test_bin_float
    output = `#{@cmd_src} -f %06.3f`
    assert_equal(0,$?.exitstatus)
    assert_equal_filecontent('expected/%s.tex' % __method__, output)
  end

  def test_bin_date
    output = `#{@cmd_src} -d "%d. %B %Y"`
    assert_equal(0,$?.exitstatus)
    assert_equal_filecontent('expected/%s.tex' % __method__, output)
  end
  
  def test_bin_boolean
    output = `#{@cmd_src} --true yes --false no`
    assert_equal(0,$?.exitstatus)
    assert_equal_filecontent('expected/%s.tex' % __method__, output)
  end 
  
  def test_bin_width10
    output = `#{@cmd_src} -w 10`
    assert_equal(0,$?.exitstatus)
    assert_equal_filecontent('expected/%s.tex' % __method__, output)
  end 
  def test_bin_width_auto
    output = `#{@cmd_src} -w auto`
    assert_equal(0,$?.exitstatus)
    assert_equal_filecontent('expected/%s.tex' % __method__, output)
  end 
end

#Make the same tests with the exe
class Test_bin_exe < Test_bin
  def setup
    super
    @cmd = '../bin/xlsx2latex.exe'
    skip('%s not available' % @cmd) unless File.exist?(@cmd)
  end
end

__END__
  