#encoding: utf-8
=begin rdoc
Tests for xlsx2latex
=end
gem 'minitest'
require 'minitest/autorun'
require 'minitest/filecontent'
  
$:.unshift('../lib')
require 'xlsx2latex'

class Test_xlsx2latex < MiniTest::Test
  def setup
    @excel = XLSX2LaTeX::Excel.new("./testdata/testdata_excel2latex.xlsx",
        log_outputter: nil
    )
  end
  def test_info
    assert_equal(<<INFO.strip, @excel.info)
File: testdata_excel2latex.xlsx
Number of sheets: 2
Sheets: Sheet1, Sheet2
Sheet 1:
  First row: 1
  Last row: 8
  First column: A
  Last column: G
Sheet 2:
  First row: 1
  Last row: 12
  First column: A
  Last column: G
INFO
  end
  def test_default_sheet
    assert_equal_filecontent('expected/%s.tex' % __method__, @excel.to_latex(sheetname: 'Sheet1'))
  end
  def test_sheet1
    assert_equal_filecontent('expected/%s.tex' % __method__, @excel.to_latex(sheetname: 'Sheet1'))
  end
  def test_sheet1_floatformat
    assert_equal_filecontent('expected/%s.tex' % __method__, @excel.to_latex(sheetname: 'Sheet1', float_format: '%0.0f',))
  end
  def test_sheet2
    assert_equal_filecontent('expected/%s.tex' % __method__, @excel.to_latex(sheetname: 'Sheet2'))
  end
end
__END__
  doc.body << 
  doc.body << excel.to_latex('Sheet2')

