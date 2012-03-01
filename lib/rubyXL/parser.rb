require 'rubygems'
require 'nokogiri'
require 'zip/zip' #rubyzip
require File.expand_path(File.join(File.dirname(__FILE__),'Hash'))

module RubyXL

  class Parser
    attr_reader :data_only, :read_only, :num_sheets
    @@parsed_column_hash ={}
    
    # converts cell string (such as "AA1") to matrix indices
    def Parser.convert_to_index(cell_string)
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
    def Parser.parse(file_path, data_only=false, read_only=false)
      @file_path = file_path
      @data_only = data_only
      @read_only = read_only

      @files = Hash.new
      Parser.decompress()
      wb = Parser.fill_workbook()
      @files = nil

      return wb
    end

    private

    #fills hashes for various styles
    def Parser.fill_styles(wb,style_hash)
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
    end

    # i is the sheet number
    # files is the hash which includes information for each worksheet
    # shared_strings has group of indexed strings which the cells reference
    def Parser.fill_worksheet(wb, i)
      sheet_names = @files['app'].css('TitlesOfParts vt|vector vt|lpstr').children
      wb.worksheets[i] = Worksheet.new(wb, sheet_names[i].to_s)

      j = i+1

      namespaces = @files[j].root.namespaces()
      unless @data_only
        sheet_views_node= @files[j].xpath('/xmlns:worksheet/xmlns:sheetViews[xmlns:sheetView]',namespaces).first
        wb.worksheets[i].sheet_view = Hash.xml_node_to_hash(sheet_views_node)[:sheetView]

        ##col styles##
        cols_node_set = @files[j].xpath('/xmlns:worksheet/xmlns:cols',namespaces)
        unless cols_node_set.empty?
          wb.worksheets[i].cols= Hash.xml_node_to_hash(cols_node_set.first)[:col]
        end
        ##end col styles##

        ##merge_cells##
        merge_cells_node = @files[j].xpath('/xmlns:worksheet/xmlns:mergeCells[xmlns:mergeCell]',namespaces)
        unless merge_cells_node.empty?
          wb.worksheets[i].merged_cells = Hash.xml_node_to_hash(merge_cells_node.first)[:mergeCell]
        end
        ##end merge_cells##

        ##sheet_view pane##
        pane_data = wb.worksheets[i].sheet_view[:pane]
        wb.worksheets[i].pane = pane_data
        ##end sheet_view pane##

        ##data_validation##
        data_validations_node = @files[j].xpath('/xmlns:worksheet/xmlns:dataValidations[xmlns:dataValidation]',namespaces)
        unless data_validations_node.empty?
          wb.worksheets[i].validations = Hash.xml_node_to_hash(data_validations_node.first)[:dataValidation]
        else
          wb.worksheets[i].validations=nil
        end
        ##end data_validation##

        #extLst
        ext_list_node = @files[j].xpath('/xmlns:worksheet/xmlns:extLst',namespaces)
        unless ext_list_node.empty?
          wb.worksheets[i].extLst = Hash.xml_node_to_hash(ext_list_node.first)
        else
          wb.worksheets[i].extLst=nil
        end
        #extLst

        ##legacy drawing##
        legacy_drawing_node = @files[j].xpath('/xmlns:worksheet/xmlns:legacyDrawing',namespaces)
        unless legacy_drawing_node.empty?
          wb.worksheets[i].legacy_drawing = Hash.xml_node_to_hash(legacy_drawing_node.first)
        else
          wb.worksheets[i].legacy_drawing = nil
        end
        ##end legacy drawing
      end

      row_data = @files[j].xpath('/xmlns:worksheet/xmlns:sheetData/xmlns:row[xmlns:c[xmlns:v]]',namespaces)
      row_data.each do |row|
        unless @data_only
          ##row styles##
          row_style = '0'
          row_attributes = row.attributes
          unless row_attributes['s'].nil?
            row_style = row_attributes['s'].value
          end

          wb.worksheets[i].row_styles[row_attributes['r'].content] = { :style => row_style  }

          if !row_attributes['ht'].nil?  && (!row_attributes['ht'].content.nil? || row_attributes['ht'].content.strip != "" )
            wb.worksheets[i].change_row_height(Integer(row_attributes['r'].content)-1,
              Float(row_attributes['ht'].content))
          end
          ##end row styles##
        end

        c_row = row.search('./xmlns:c')
        c_row.each do |value|
          value_attributes= value.attributes
          cell_index = Parser.convert_to_index(value_attributes['r'].content)
          style_index = nil

          data_type = value_attributes['t'].content if value_attributes['t']
          element_hash ={}
          value.children.each do |node|
            element_hash["#{node.name()}_element"]=node
          end
          # v is the value element that is part of the cell
          if element_hash["v_element"]
            v_element_content = element_hash["v_element"].content
          else
            v_element_content=""
          end
          if v_element_content =="" #no data
            cell_data = nil
          elsif data_type == 's' #shared string
            str_index = Integer(v_element_content)
            cell_data = wb.shared_strings[str_index].to_s
          elsif data_type=='str' #raw string
            cell_data = v_element_content
          elsif data_type=='e' #error
            cell_data = v_element_content
          else# (value.css('v').to_s != "") && (value.css('v').children.to_s != "") #is number
            data_type = ''
            if(v_element_content =~ /\./) #is float
              cell_data = Float(v_element_content)
            else
              cell_data = Integer(v_element_content)
            end
          end
          cell_formula = nil
          fmla_css = element_hash["f_element"]
          if fmla_css && fmla_css.content
            fmla_css_content= fmla_css.content
            if(fmla_css_content != "")
              cell_formula = fmla_css_content
              cell_formula_attr = {}
              fmla_css_attributes = fmla_css.attributes
              cell_formula_attr['t'] = fmla_css_attributes['t'].content if fmla_css_attributes['t']
              cell_formula_attr['ref'] = fmla_css_attributes['ref'].content if fmla_css_attributes['ref']
              cell_formula_attr['si'] = fmla_css_attributes['si'].content if fmla_css_attributes['si']
            end
          end

          unless @data_only
            style_index = value['s'].to_i #nil goes to 0 (default)
          else
            style_index = 0
          end

          wb.worksheets[i].sheet_data[cell_index[0]] ||= Array.new
          wb.worksheets[i].sheet_data[cell_index[0]][cell_index[1]] =
            Cell.new(wb.worksheets[i],cell_index[0],cell_index[1],cell_data,cell_formula,
              data_type,style_index,cell_formula_attr)
        end
      end
    end
    
    def Parser.parse_xml(name, path)
      # figure out parse options
      parse_options = Nokogiri::XML::ParseOptions::DEFAULT_XML
      parse_options |= Nokogiri::XML::ParseOptions::COMPACT if @read_only

      # Open, parse, and store it
      if File.exist?(path)
        File.open(path, 'rb') do |f|
          @files[name] = Nokogiri::XML.parse(f)
        end
      end
    end

    def Parser.read_external_files(path)
      retval = nil
      
      if File.directory?(path)
        retval = {}
        entries = Dir.new(path).entries.reject { |f| File.directory?(File.join(path, f)) || f == ".DS_Store" }
        entries.each_with_index do |filename, i|
          File.open(File.join(path, filename), 'rb') do |f|
            retval[i+1] = f.read
          end
        end
      else
        File.open(path, 'rb') do |f|
          retval = f.read
        end
      end
      
      retval
    end

    def Parser.decompress()
      # ensure it is an xlsx/xlsm file
      if @file_path =~ /(.+)\.xls(x|m)/
        dir_path = $1.to_s
      else
        raise 'Not .xlsx or .xlsm excel file'
      end

      # copy excel file to zip file in same directory
      dir_path = File.join(File.dirname(dir_path), make_safe_name(Time.now.to_s))
      zip_path = dir_path + '.zip'
      FileUtils.cp(@file_path,zip_path)
      MyZip.new.unzip(zip_path,dir_path)
      File.delete(zip_path)

      # parse core files
      Parser.parse_xml('app', File.join(dir_path, 'docProps', 'app.xml'))
      Parser.parse_xml('core', File.join(dir_path, 'docProps', 'core.xml'))
      Parser.parse_xml('workbook', File.join(dir_path, 'xl', 'workbook.xml'))
      Parser.parse_xml('sharedString', File.join(dir_path, 'xl', 'sharedStrings.xml'))

      # preserve external links
      unless @data_only
        @files['externalLinks'] = Parser.read_external_files(File.join(dir_path, 'xl', 'externalLinks'))
        @files['externalLinks']['rels'] = Parser.read_external_files(File.join(dir_path, 'xl', 'externalLinks', '_rels'))
        @files['drawings'] = Parser.read_external_files(File.join(dir_path, 'xl', 'drawings'))
        @files['printerSettings'] = Parser.read_external_files(File.join(dir_path, 'xl', 'printerSettings'))
        @files['worksheetRels'] = Parser.read_external_files(File.join(dir_path, 'xl', 'worksheets', '_rels'))
        @files['vbaProject'] = Parser.read_external_files(File.join(dir_path, 'xl', 'vbaProject.bin'))
        Parser.parse_xml('styles', File.join(dir_path, 'xl', 'styles.xml'))
      end

      # parse the worksheets
      @num_sheets = Integer(@files['workbook'].css('sheets').children.size)
      for i in 1..@num_sheets
        filename = 'sheet' + i.to_s + '.xml'
        Parser.parse_xml(i, File.join(dir_path, 'xl', 'worksheets', filename))
      end

      # cleanup
      FileUtils.rm_rf(dir_path)
    end

    def Parser.fill_workbook()
      wb = Workbook.new([nil],@file_path)
      wb.worksheets = Array.new(@num_sheets) # array of Worksheet objs

      #attr_accessor :worksheets, :filepath, :creator, :modifier, :created_at,
      #  :modified_at, :company, :application, :appversion, :num_fmts, :fonts, :fills,
      #  :borders, :cell_xfs, :cell_style_xfs, :cell_styles, :shared_strings, :calc_chain,
      #  :num_strings, :size, :date1904, :external_links, :style_corrector, :drawings,
      #  :worksheet_rels, :printer_settings, :macros, :colors, :shared_strings_XML, :defined_names, :column_lookup_hash

      unless @data_only
        wb.creator = @files['core'].css('dc|creator').children.to_s
        wb.modifier = @files['core'].css('cp|last_modified_by').children.to_s
        wb.created_at = @files['core'].css('dcterms|created').children.to_s
        wb.modified_at = @files['core'].css('dcterms|modified').children.to_s

        wb.company = @files['app'].css('Company').children.to_s
        wb.application = @files['app'].css('Application').children.to_s
        wb.appversion = @files['app'].css('AppVersion').children.to_s
      end

      wb.defined_names = @files['workbook'].css('definedNames').to_s
      wb.date1904 = @files['workbook'].css('workbookPr').attribute('date1904').to_s == '1'

      wb.shared_strings = {}
      unless @files['sharedString'].nil?
        wb.shared_strings_XML = @files['sharedString'].to_s unless @read_only
        wb.num_strings = Integer(@files['sharedString'].css('sst').attribute('count').value())
        wb.size = Integer(@files['sharedString'].css('sst').attribute('uniqueCount').value())

        @files['sharedString'].css('si').each do |node|
          unless node.css('r').empty?
            text = node.css('r t').children.to_a.join
            node.children.remove
            node << "<t xml:space=\"preserve\">#{text}</t>"
          end
        end

        @files['sharedString'].css('si t').each_with_index do |node, i|
          string = node.children.to_s
          wb.shared_strings[i] = string
          wb.shared_strings[string] = i unless @read_only
        end
      end

      unless @data_only
        style_hash = Hash.xml_node_to_hash(@files['styles'].root)
        Parser.fill_styles(wb,style_hash)

        # will be nil if these files do not exist
        wb.external_links = @files['externalLinks']
        wb.drawings = @files['drawings']
        wb.printer_settings = @files['printerSettings']
        wb.worksheet_rels = @files['worksheetRels']
        wb.macros = @files['vbaProject']
      end

      # for each worksheet:
      # 1. find the dimensions of the data matrix
      # 2. Fill in the matrix with data from worksheet/shared_string files
      # 3. Apply styles
      for i in 0..@num_sheets-1
        Parser.fill_worksheet(wb, i)
      end

      wb
    end

    def Parser.safe_filename(name, allow_mb_chars=false)
      # "\w" represents [0-9A-Za-z_] plus any multi-byte char
      regexp = allow_mb_chars ? /[^\w]/ : /[^0-9a-zA-Z\_]/
      name.gsub(regexp, "_")
    end

    # Turns the passed in string into something safe for a filename
    def Parser.make_safe_name(name, allow_mb_chars=false)
      ext = safe_filename(File.extname(name), allow_mb_chars).gsub(/^_/, '.')
      "#{safe_filename(name.gsub(ext, ""), allow_mb_chars)}#{ext}".gsub(/\(/, '_').gsub(/\)/, '_').gsub(/__+/, '_').gsub(/^_/, '').gsub(/_$/, '')
    end

  end
end
