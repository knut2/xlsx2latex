#encoding: utf-8
=begin rdoc
Tests for xlsx2latex
=end
gem 'minitest'
require 'minitest/autorun'
require 'minitest/filecontent'
  
$:.unshift('../lib')
require 'xlsx2latex'

class Test_generated_xlsx < MiniTest::Test
  def setup
    @excel = XLSX2LaTeX::Excel.new("./testdata/generated.xlsx",
        log_outputter: nil
    )
  end
  def test_info
#~ puts @excel.info    
    assert_equal(<<INFO.strip, @excel.info)
File: generated.xlsx
Number of sheets: 3
Sheets: DATA, Merged Areas, Formats
Sheet 1:
  First row: 1
  Last row: 7
  First column: A
  Last column: H
Sheet 2:
  First row: 1
  Last row: 15
  First column: A
  Last column: K
Sheet 3:
  First row: 1
  Last row: 5
  First column: A
  Last column: A
INFO
  end
  def test_default_sheet
    assert_equal_filecontent('expected/generated_%s.tex' % __method__, @excel.to_latex())
  end
  def test_sheet_data
    assert_equal_filecontent('expected/generated_%s.tex' % __method__, @excel.to_latex(sheetname: 'DATA'))
  end
  def test_sheet_data_stringformat_10
    assert_equal_filecontent('expected/generated_%s.tex' % __method__, @excel.to_latex(sheetname: 'DATA', string_format: '%-10s'))
  end
  def test_sheet_data_floatformat
    assert_equal_filecontent('expected/generated_%s.tex' % __method__, @excel.to_latex(sheetname: 'DATA', float_format: '%0.0f',))
  end
  def test_sheet_data_date_format
    assert_equal_filecontent('expected/generated_%s.tex' % __method__, @excel.to_latex(sheetname: 'DATA', date_format: '%d. %B %Y',))
  end
  def test_sheet_data_time_format
    assert_equal_filecontent('expected/generated_%s.tex' % __method__, @excel.to_latex(sheetname: 'DATA', time_format: '%d. %B %Y %H:%M',))
  end
  def test_sheet_data_boolean
    assert_equal_filecontent('expected/generated_%s.tex' % __method__, @excel.to_latex(sheetname: 'DATA', bool_true: 'yes', bool_false: 'no'))
  end
  def test_sheet_merged_areas
    assert_equal_filecontent('expected/generated_%s.tex' % __method__, @excel.to_latex(sheetname: 'Merged Areas'))
  end
  def test_sheet_format
    assert_equal_filecontent('expected/generated_%s.tex' % __method__, @excel.to_latex(sheetname: 'Formats'))
  end
end
__END__
  doc.body << 
  doc.body << excel.to_latex('Sheet2')

