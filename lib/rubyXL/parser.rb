require 'rubygems'
require 'nokogiri'
require 'zip/zip'
require 'date'
require File.expand_path(File.join(File.dirname(__FILE__),'Hash'))

module RubyXL
  
  # Slightly modified from https://gist.github.com/827475
  # A small DSL for helping parsing documents using Nokogiri::XML::Reader. The
  # XML Reader is a good way to move a cursor through a (large) XML document fast,
  # but is not as cumbersome as writing a full SAX document handler. Read about
  # it here: http://nokogiri.org/Nokogiri/XML/Reader.html
  # 
  # Just pass the reader in this parser and specificy the nodes that you are interested
  # in in a block. You can just parse every node or only look inside certain nodes.
  # 
  # A small example:
  # 
  # Reader.new(filename) do
  #   inside_element 'User' do
  #     for_element 'Name' do puts "Username: #{inner_xml}" end
  #     for_element 'Email' do puts "Email: #{inner_xml}" end
  #     
  #     for_element 'Address' do
  #       puts 'Start of address:'
  #       inside_element do
  #         for_element 'Street' do puts "Street: #{inner_xml}" end
  #         for_element 'Zipcode' do puts "Zipcode: #{inner_xml}" end
  #         for_element 'City' do puts "City: #{inner_xml}" end
  #       end
  #       puts 'End of address'
  #     end
  #   end
  # end
  # 
  # It does NOT fail on missing tags, and does not guarantee order of execution. It parses
  # every tag regardless of nesting. The only way to guarantee scope is by using
  # the `inside_element` method. This limits the parsing to the current or the named tag.
  # If tags are encountered multiple times, their blocks will be called multiple times.
  class Reader
    def initialize(filename, &block)
      @node = Nokogiri::XML::Reader(File.new(filename))
      @node.each do
        self.instance_eval &block
      end
    end

    def name
      @node.name
    end

    def inner_xml
      @node.inner_xml.strip
    end

    def is_start?
      @node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT
    end

    def is_end?
      @node.node_type == Nokogiri::XML::Reader::TYPE_END_ELEMENT
    end

    def attribute(attribute)
      @node.attribute(attribute)
    end

    def for_element(name, &block)
      return unless self.name == name and is_start?
      self.instance_eval &block
    end

    def inside_element(name=nil, &block)
      return if @node.self_closing?
      return unless name.nil? or (self.name == name and is_start?)

      name = @node.name
      depth = @node.depth

      @node.each do
        return if self.name == name and is_end? and @node.depth == depth
        self.instance_eval &block
      end
    end
  end

  class Parser
    @@parsed_column_hash = {}
    @@debug = true
    
    # converts cell string (such as "AA1") to matrix indices
    def self.convert_to_index(cell_string)
      index = Array.new(2)
      index[0]=-1
      index[1]=-1
      if(cell_string =~ /^([A-Z]+)(\d+)$/)

        one = $1
        row = $2.to_i - 1 #-1 for 0 indexing
        col = 0
        i = 0
        if @@parsed_column_hash[one].nil?
          two = one.reverse #because of 26^i calculation
          two.each_byte do |c|
            int_val = c - 64 #converts A to 1
            col += int_val * 26**(i)
            i=i+1
          end
          @@parsed_column_hash[one] = col
        else
          col = @@parsed_column_hash[one]
        end
        col -= 1 #zer0 index
        index[0] = row
        index[1] = col
      end
      return index
    end

    # data_only allows only the sheet data to be parsed, so as to speed up parsing
    #   However, using this option will result in date-formatted cells being interpreted as numbers
    # read_only disables modification or writing of the file, but results in a much
    #   lower memory footprint
    def self.parse(file_path, data_only=false, read_only=false)
      # ensure we are given a xlsx/xlsm file
      if file_path =~ /(.+)\.xls(x|m)/
        dir_path = $1.to_s
      else
        raise 'Not .xlsx or .xlsm excel file'
      end

      @data_only = data_only
      @read_only = read_only

      # copy excel file to zip file in same directory
      puts "[#{Time.now}] Uncompressing #{file_path}..." if @@debug
      dir_path = File.join(File.dirname(dir_path), make_safe_name(Time.now.to_s))
      zip_path = dir_path + '.zip'
      FileUtils.cp(file_path,zip_path)
      MyZip.new.unzip(zip_path,dir_path)
      File.delete(zip_path)
      puts "[#{Time.now}] done." if @@debug

      # parse workbook.xml
      workbook_xml = Parser.parse_xml(File.join(dir_path, 'xl', 'workbook.xml'))
      
      # build workbook
      @num_sheets = Integer(workbook_xml.css('sheets').children.size)
      wb = Workbook.new([nil], file_path)
      wb.worksheets = Array.new(@num_sheets)

      # extract everything we need from workbook.xml
      wb.defined_names = workbook_xml.css('definedNames').to_s
      wb.date1904 = workbook_xml.css('workbookPr').attribute('date1904').to_s == '1'
      workbook_xml = nil

      # extract everything we need from app.xml
      app_xml = Parser.parse_xml(File.join(dir_path, 'docProps', 'app.xml'))
      sheet_names = app_xml.css('TitlesOfParts vt|vector vt|lpstr').children
      unless @data_only
        wb.company = app_xml.css('Company').children.to_s
        wb.application = app_xml.css('Application').children.to_s
        wb.appversion = app_xml.css('AppVersion').children.to_s
      end
      app_xml = nil
      
      # extract everything we need from core.xml
      unless @data_only
        core_xml = Parser.parse_xml(File.join(dir_path, 'docProps', 'core.xml'))
        wb.creator = core_xml.css('dc|creator').children.to_s
        wb.modifier = core_xml.css('cp|last_modified_by').children.to_s
        wb.created_at = core_xml.css('dcterms|created').children.to_s
        wb.modified_at = core_xml.css('dcterms|modified').children.to_s
        core_xml = nil
      end

      # extract everything we need from sharedStrings.xml
      wb.shared_strings = {}
      shared_strings_xml = Parser.parse_xml(File.join(dir_path, 'xl', 'sharedStrings.xml'))
      unless shared_strings_xml.nil?
        puts "[#{Time.now}] Processing shared strings (phase 1)..." if @@debug
        wb.shared_strings_XML = shared_strings_xml.to_s unless @read_only
        wb.num_strings = Integer(shared_strings_xml.css('sst').attribute('count').value())
        wb.size = Integer(shared_strings_xml.css('sst').attribute('uniqueCount').value())
        puts "[#{Time.now}] done." if @@debug

        puts "[#{Time.now}] Processing shared strings (phase 2)..." if @@debug
        shared_strings_xml.css('si').each_with_index do |node, i|
          # Merge si > r > t into si > t; unnecessary?
          #unless node.css('r').empty?
          #  text = node.css('r t').children.to_a.join
          #  node.children.remove
          #  node << "<t xml:space=\"preserve\">#{text}</t>"
          #end

          # Build two way hash for shared strings
          node.css('t').each do |t|
            t = t.child
            unless t.nil?
              str = t.content
              wb.shared_strings[i] = str
              wb.shared_strings[str] = i unless @read_only
            end
          end
        end
        puts "[#{Time.now}] done." if @@debug

        shared_strings_xml = nil
      end

      # preserve external links
      unless @data_only
        wb.external_links = Parser.read_external_files(File.join(dir_path, 'xl', 'externalLinks'))
        wb.drawings = Parser.read_external_files(File.join(dir_path, 'xl', 'drawings'))
        wb.printer_settings = Parser.read_external_files(File.join(dir_path, 'xl', 'printerSettings'))
        wb.worksheet_rels = Parser.read_external_files(File.join(dir_path, 'xl', 'worksheets', '_rels'))
        wb.macros = Parser.read_external_files(File.join(dir_path, 'xl', 'vbaProject.bin'))
        
        styles_xml = Parser.parse_xml(File.join(dir_path, 'xl', 'styles.xml'))
        Parser.fill_styles(wb, Hash.xml_node_to_hash(styles_xml.root))
        styles_xml = nil
      end

      # parse the worksheets
      for i in 0..@num_sheets-1
        filename = 'sheet' + (i+1).to_s + '.xml'
        wb.worksheets[i] = Parser.parse_worksheet(wb, File.join(dir_path, 'xl', 'worksheets', filename))
        wb.worksheets[i].sheet_name = sheet_names[i].to_s
      end

      # cleanup
      FileUtils.rm_rf(dir_path)
      return wb
    end

    private

    # fill hashes for various styles
    def self.fill_styles(wb,style_hash)
      puts "[#{Time.now}] Filling styles..." if @@debug
      
      wb.num_fmts = style_hash[:numFmts]

      ###FONTS###
      wb.fonts = {}
      if style_hash[:fonts][:attributes][:count]==1
        style_hash[:fonts][:font] = [style_hash[:fonts][:font]]
      end

      style_hash[:fonts][:font].each_with_index do |f,i|
        wb.fonts[i.to_s] = {:font=>f,:count=>0}
      end

      ###FILLS###
      wb.fills = {}
      if style_hash[:fills][:attributes][:count]==1
        style_hash[:fills][:fill] = [style_hash[:fills][:fill]]
      end

      style_hash[:fills][:fill].each_with_index do |f,i|
        wb.fills[i.to_s] = {:fill=>f,:count=>0}
      end

      ###BORDERS###
      wb.borders = {}
      if style_hash[:borders][:attributes][:count] == 1
        style_hash[:borders][:border] = [style_hash[:borders][:border]]
      end

      style_hash[:borders][:border].each_with_index do |b,i|
        wb.borders[i.to_s] = {:border=>b, :count=>0}
      end

      wb.cell_style_xfs = style_hash[:cellStyleXfs]
      wb.cell_xfs = style_hash[:cellXfs]
      wb.cell_styles = style_hash[:cellStyles]

      wb.colors = style_hash[:colors]

      #fills out count information for each font, fill, and border
      if wb.cell_xfs[:xf].is_a?(::Hash)
        wb.cell_xfs[:xf] = [wb.cell_xfs[:xf]]
      end
      wb.cell_xfs[:xf].each do |style|
        id = style[:attributes][:fontId].to_s
        unless id.nil?
          wb.fonts[id][:count] += 1
        end

        id = style[:attributes][:fillId].to_s
        unless id.nil?
          wb.fills[id][:count] += 1
        end

        id = style[:attributes][:borderId].to_s
        unless id.nil?
          wb.borders[id][:count] += 1
        end
      end

      puts "[#{Time.now}] done." if @@debug
    end

    # parse worksheet
    def self.parse_worksheet(wb, filename)
      worksheet = Worksheet.new(wb)
      
      puts "[#{Time.now}] Parsing '#{filename}'..." if @@debug
      Reader.new(filename) do
        inside_element 'worksheet' do
          inside_element 'sheetData' do
            inside_element 'row' do
              unless @data_only
                ##row styles##
                if is_start?
                  row_style = '0'
                  row_style = attribute('s') unless attribute('s').nil?

                  worksheet.row_styles[attribute('r')] = { :style => row_style  }

                  if !attribute('ht').nil? && !attribute('ht').strip == ""
                    worksheet.change_row_height(Integer(attribute('r'))-1, Float(attribute('ht')))
                  end
                end
                ##end row styles##
              end

              inside_element 'c' do
                # Scan attributes
                cell_index = Parser.convert_to_index(attribute('r'))
                data_type = attribute('t')
              
                # Get cell data and coerce type
                cell_data = nil
                for_element 'v' do
                  if data_type == 's' # shared string
                    cell_data = worksheet.workbook.shared_strings[Integer(value)]
                  elsif data_type == 'str' # raw string
                    cell_data = value
                  elsif data_type == 'e' # error
                    cell_data = value
                  else # (value.css('v').to_s != "") && (value.css('v').children.to_s != "") #is number
                    data_type = ''
                    if value =~ /\./ #is float
                      cell_data = Float(value)
                    else
                      cell_data = Integer(value)
                    end
                  end
                end
              
                # Parse out formula
                cell_formula = nil
                cell_formula_attr = {}
                for_element 'f' do
                  if !value.nil? && value != ''
                    cell_formula = value
                    cell_formula_attr['t'] = attribute('t')
                    cell_formula_attr['ref'] = attribute('ref')
                    cell_formula_attr['si'] = attribute('si')
                  end
                end

                # Get style
                style_index = 0
                for_element 's' do
                  unless @data_only
                    style_index = value.to_i # nil goes to 0 (default)
                  end
                end

                # Add Cell
                worksheet.sheet_data[cell_index[0]][cell_index[1]] = Cell.new(worksheet, cell_index[0], cell_index[1],
                  cell_data, cell_formula, data_type, style_index, cell_formula_attr)
              end
            end
          end
        end
      end
      puts "[#{Time.now}] done." if @@debug

      worksheet_xml = Parser.parse_xml(filename)
      unless @data_only
        sheet_views_node = worksheet_xml.xpath('/xmlns:worksheet/xmlns:sheetViews[xmlns:sheetView]').first
        worksheet.sheet_view = Hash.xml_node_to_hash(sheet_views_node)[:sheetView]
        sheet_views_node = nil

        ##col styles##
        cols_node_set = worksheet_xml.xpath('/xmlns:worksheet/xmlns:cols')
        unless cols_node_set.empty?
          worksheet.cols = Hash.xml_node_to_hash(cols_node_set.first)[:col]
        end
        cols_node_set = nil
        ##end col styles##

        ##merge_cells##
        merge_cells_node = worksheet_xml.xpath('/xmlns:worksheet/xmlns:mergeCells[xmlns:mergeCell]')
        unless merge_cells_node.empty?
          worksheet.merged_cells = Hash.xml_node_to_hash(merge_cells_node.first)[:mergeCell]
        end
        merge_cells_node = nil
        ##end merge_cells##

        ##sheet_view pane##
        worksheet.pane = worksheet.sheet_view[:pane]
        ##end sheet_view pane##

        ##data_validation##
        data_validations_node = worksheet_xml.xpath('/xmlns:worksheet/xmlns:dataValidations[xmlns:dataValidation]')
        worksheet.validations = nil
        unless data_validations_node.empty?
          worksheet.validations = Hash.xml_node_to_hash(data_validations_node.first)[:dataValidation]
        end
        data_validations_node = nil
        ##end data_validation##

        #extLst
        ext_list_node = worksheet_xml.xpath('/xmlns:worksheet/xmlns:extLst')
        worksheet.extLst = nil
        unless ext_list_node.empty?
          worksheet.extLst = Hash.xml_node_to_hash(ext_list_node.first)
        end
        ext_list_node = nil
        #extLst

        ##legacy drawing##
        legacy_drawing_node = worksheet_xml.xpath('/xmlns:worksheet/xmlns:legacyDrawing')
        worksheet.legacy_drawing = nil
        unless legacy_drawing_node.empty?
          worksheet.legacy_drawing = Hash.xml_node_to_hash(legacy_drawing_node.first)
        end
        legacy_drawing_node = nil
        ##end legacy drawing
      end

      worksheet
    end
    
    def self.parse_xml(path)
      # figure out parse options
      parse_options = Nokogiri::XML::ParseOptions::DEFAULT_XML
      parse_options |= Nokogiri::XML::ParseOptions::COMPACT if @read_only

      retval = nil

      # Open, parse, and store it
      if File.exist?(path)
        puts "[#{Time.now}] Parsing #{path}..." if @@debug
        File.open(path, 'rb') do |f|
          retval = Nokogiri::XML.parse(f, nil, nil, parse_options)
        end
        puts "[#{Time.now}] done." if @@debug
      end

      retval
    end

    def self.read_external_files(path)
      retval = nil
      
      if File.directory?(path)
        retval = {}
        entries = Dir.new(path).entries.reject { |f| File.directory?(File.join(path, f)) || f == ".DS_Store" }
        entries.each_with_index do |filename, i|
          puts "[#{Time.now}] Reading #{path}..." if @@debug
          File.open(File.join(path, filename), 'rb') do |f|
            retval[i+1] = f.read
          end
          puts "[#{Time.now}] done." if @@debug
        end
      else
        puts "[#{Time.now}] Reading #{path}..." if @@debug
        File.open(path, 'rb') do |f|
          retval = f.read
        end
        puts "[#{Time.now}] done." if @@debug
      end
      
      retval
    end

    def self.safe_filename(name, allow_mb_chars=false)
      # "\w" represents [0-9A-Za-z_] plus any multi-byte char
      regexp = allow_mb_chars ? /[^\w]/ : /[^0-9a-zA-Z\_]/
      name.gsub(regexp, "_")
    end

    # Turns the passed in string into something safe for a filename
    def self.make_safe_name(name, allow_mb_chars=false)
      ext = safe_filename(File.extname(name), allow_mb_chars).gsub(/^_/, '.')
      "#{safe_filename(name.gsub(ext, ""), allow_mb_chars)}#{ext}".gsub(/\(/, '_').gsub(/\)/, '_').gsub(/__+/, '_').gsub(/^_/, '').gsub(/_$/, '')
    end

  end
end
