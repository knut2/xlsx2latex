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
* string_format: '%s'
* float_format: '%0.2f'
* date_format: '%Y-%m-%d',
* time_format: '%Y-%m-%d %H:%M',
* Boolean:
  * bool_true: 'X', 
  * bool_false: '-',
* auto_width: See details below

==auto_width
The option quto_width sets the column width to longest value of the column.

* The columns are right justified.
  (On this level there is no difference between numbers and strings.
  And it is easier to read right justified strings then left justified numbers).
* In combination with multicolumns the result may be confusing.
=end
    def initialize(filename,
              string_format: '%s',
              float_format: '%0.2f',
              date_format: '%Y-%m-%d',
              time_format: '%Y-%m-%d %H:%M',
              bool_true: 'X', bool_false: '-',
              auto_width: false,
              log_outputter: Log4r::StdoutOutputter.new('stdout')
            )
      @filename = filename
      @log = Log4r::Logger.new(File.basename(filename))
      @log.outputters << log_outputter if log_outputter
      @log.level = Log4r::INFO
      
      @xlsx = Roo::Excelx.new(filename)      
      @string_format = string_format
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
                  string_format: @string_format,
                  float_format: @float_format,
                  date_format: @date_format,
                  time_format: @time_format,
                  bool_true: @bool_true, bool_false: @bool_false,
                  auto_width: @auto_width,
                  xxx: nil
        )
      @xlsx.sheet(sheetname)  #set default sheetname
      sheet = @xlsx.sheet_for(sheetname)
      @log.info("Convert Sheet %s to LaTeX" % sheetname)
      
      tex = []
      max_columns = 1
      widths = Hash.new(0) #maximum width per column. Needed for auto_width
      sheet.each_row do |line|
        lineval = []
        max_columns = @xlsx.last_column if @xlsx.last_column > max_columns
        line.each_with_index{| cell,colnum |
          @log.debug("%s(%i,%i): %s <%s> %s" % [sheetname,cell.coordinate.column,cell.coordinate.row, cell.type.inspect, cell.value.inspect, cell.inspect]) if @log.debug?
          right_justified = false
        
          case cell.type
          when nil  #Overwritten parts in multicol/multirow
            @log.warn("%s(%i,%i) contains Nil - Ok for multicolumn. Multirow is not supported" % [sheetname,cell.coordinate.column,cell.coordinate.row])
            case lineval.last
              when /multicolumn\{(\d+)\}/
                lineval.last.sub!(/multicolumn\{(\d+)\}/, 'multicolumn{%i}' % [$1.to_i + 1])
              else #First combined cell
                #Ignore the width of the mutlicolumn command.
                #I expect the merged cells are long enough to compensate the additional length
                widths[colnum-1] = lineval.last.size if lineval.last.size > widths[colnum-1]
                lineval << ('\multicolumn{2}{c}{%s}' % lineval.pop)
              end
            next
          when :float
            value = float_format % cell.value
          when :string
            value = string_format % cell.value
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
            value = '\textbf{%s}' % value.strip if font.bold?
            value = '\emph{%s}'  % value.strip if font.italic?
            value = '\underline{%s}' % value.strip if font.underline?
            value = string_format % value if cell.type == :string
          end
          lineval  << value
          widths[colnum] = value.to_s.size if value.to_s.size > widths[colnum]
        }
        tex << lineval
      end
      #Prepare auto width if requested
      tex.map{|tabline|
        tabline.map!.with_index{|cell,i| 
          if cell =~ /\\multicolumn\{(\d+)\}/ #Calculate offset for next column
            width = -3  #3 corresponds to the ' & ' and we need one less.
            i.upto($1.to_i){|i2| width += widths[i2] + 3} #calculate width up to next cell
            '%-*s' % [width,cell]
          else
            '%*s' % [widths[i],cell]
          end
        }
      } if auto_width

      [
        "\\begin{tabular}{%s}" % ('c|'* max_columns),
        tex.map{|tabline| tabline.join(' & ')}.join("\\\\\n"),
        '\end{tabular}'
      ].join("\n")
    end
  end
end #XLSX2LaTeX
