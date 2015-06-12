#encoding: utf-8
=begin rdoc
=end
module XLSX2LaTeX
=begin rdoc
Read an Excel-file (xlsx).

For conversion rules see #to_latex.

Limits:
* No support of multirow

Not realized yet:
* no lines, border
=end
  class Excel
=begin rdoc
You can set the default values for the conversion rules in #to_latex.

call-seq:
  Excel.new(filename, [options])

The supported options (with defaults):
* float_format: '%0.2f'
* date_format: '%Y-%m-%d',
* time_format: '%Y-%m-%d %H:%M',
* Boolean:
  * bool_true: 'X', 
  * bool_false: '-',

=end
    def initialize(filename,
              float_format: '%0.2f',
              date_format: '%Y-%m-%d',
              time_format: '%Y-%m-%d %H:%M',
              bool_true: 'X', bool_false: '-',
              log_outputter: Log4r::StdoutOutputter.new('stdout')
            )
      @filename = filename
      @log = Log4r::Logger.new(File.basename(filename))
      @log.outputters << log_outputter if log_outputter
      @log.level = Log4r::INFO
      
      @xlsx = Roo::Excelx.new(filename)      
      @float_format = float_format
      @date_format = date_format
      @time_format = time_format
      @bool_true = bool_true
      @bool_false = bool_false
      
    end
    #Logger.
    attr_reader :log
    #The internal Roo::Excelx-instance.
    attr_reader :xlsx
=begin rdoc
Call #xlsx.info
=end
    def info;  @xlsx.info; end
    def inspect;  '<XLSX2LaTeX::Excel %s>' % @filename; end
=begin rdoc

call-seq:
  #to_latex(options)
  
Options are:
* sheetname ( #xlsx#default_sheet )
* The default from ::new for the conversions can be redefined for each tabular.
=end
    def to_latex(sheetname: @xlsx.default_sheet,
                  float_format: @float_format,
                  date_format: @date_format,
                  time_format: @time_format,
                  bool_true: @bool_true, bool_false: @bool_false,
                  xxx: nil
        )
      @xlsx.sheet(sheetname)  #set default sheetname
      sheet = @xlsx.sheet_for(sheetname)
      @log.info("Convert Sheet %s to LaTeX" % sheetname)

      tex = []
      max_columns = 1
      sheet.each_row do |line|
        lineval = []
        max_columns = @xlsx.last_column if @xlsx.last_column > max_columns
        line.each{| cell |
          @log.debug("%s(%i,%i): %s <%s> %s" % [sheetname,cell.coordinate.column,cell.coordinate.row, cell.type.inspect, cell.value.inspect, cell.inspect]) if @log.debug?
        
          case cell.type
          when nil  #Overwritten parts in multicol/multirow
            @log.warn("%s(%i,%i) contains Nil - Ok for multicolumn. Multirow is not supported" % [sheetname,cell.coordinate.column,cell.coordinate.row])
            case lineval.last
              when /multicolumn\{(\d+)\}/
                lineval.last.sub!(/multicolumn\{(\d+)\}/, 'multicolumn{%i}' % [$1.to_i + 1])
              else #First combined cell
                lineval << ('\multicolumn{2}{c}{%s}' % lineval.pop)
              end
            next
          when :float
            value = float_format % cell.value
          when :string
            value = cell.value
          when :formula
            @log.info("%s(%i,%i): Found formula %s. Take result %s" % [sheetname,cell.coordinate.column,cell.coordinate.row, cell.formula, cell.value])
            value = cell.value
          when :boolean
            value = case cell.value
              when 'TRUE'; bool_true
              when 'FALSE'; bool_false
              else
                @log.warn("Don't know how to handle boolean %s" % cell.value)
                cell.value
            end
          when :date
            value = cell.value.strftime(date_format)
          when :datetime
            value = cell.value.strftime(time_format)
          else
          #~  * :date * :percentage * :time * :datetime.
            @log.error("%s(%i,%i) Undefined celltype %s -%s" % [sheetname,cell.coordinate.column,cell.coordinate.row,cell.type.inspect, cell.inspect])
            #~ value = cell.excelx_value
            value = cell.value
          end
          if value.respond_to?(:gsub)
            value.gsub!('<', '\textless{}')
            value.gsub!('>', '\textgreater{}')
            value.gsub!('$', '\$')
            value.gsub!('&', '\&')
            value.gsub!('%', '\%')
          end
          
          if font = @xlsx.font(cell.coordinate.row,cell.coordinate.column)
            value = '\textbf{%s}' % value if font.bold?
            value = '\emph{%s}'  % value if font.italic?
            value = '\underline{%s}' % value if font.underline?
          end
          lineval  << value
        }
        tex << lineval.join(' & ')
      end
      [
        "\\begin{tabular}{%s}" % ('c|'* max_columns),
        tex.join("\\\\\n"),
        '\end{tabular}'
      ].join("\n")
    end
  end
end #XLSX2LaTeX
