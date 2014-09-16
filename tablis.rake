# -*- coding: utf-8 -*-

module Tablis
  class << self
    def load(*files)
      DataModel.new(*files)
    end
  end

  module Common
    def replace_file(file, lines)
      File.rename(file, "#{file}.bak")
      File.open(file, "w") {|f| f.puts lines }
    end

    def index_and_delete(lines, regex)
      i = lines.index {|line| line =~ regex }
      lines.delete_if {|line| line =~ regex }
      i
    end

    def comment(*suffixes)
      "# TB:" + suffixes.join(':')
    end

    def regex(*suffixes)
      Regexp.new(comment(*suffixes))
    end

    def to_hankaku(str)
      str.tr('Ａ-Ｚａ-ｚ１-９＃♯＿', 'A-Za-z1-9##_')
    end
  end

  class DataModel
    include Common

    # 0: symbol
    # 1: header field name
    # 2: if the value type is boolean
    # 3: if the header field name is required
    # 4: if the value is required
    TABLE_LIST_FIELDS = [
      [:comment, '#', true],
      [:logical_name, '論理テーブル名', false, true, true],
      [:model_name, 'モデル名', false, true, true],
      [:table_name, '物理テーブル名'],
      [:no_primary_key, '主キーなし', true],
      [:primary_key, '物理主キー名'],
    ]

    RELATION_FIELDS = [
      [:comment, '#', true],
      [:table1, 'テーブル1', false, true, true],
      [:table2, 'テーブル2', false, true, true],
      [:foreign1, '外部キー1', false, true],
      [:foreign2, '外部キー2', false, true],
      [:default1, 'デフォルト1'],
      [:default2, 'デフォルト2'],
      [:index1, 'インデックス1', true],
      [:index2, 'インデックス2', true],
      [:option1, 'オプション1'],
      [:option2, 'オプション2'],
      [:rel_option1, '関係オプション1'],
      [:rel_option2, '関係オプション2'],
      [:relation, '関係', false, true, true],
      [:through, 'スルー'],
    ]

    COLUMN_FIELDS = [
      [:comment, '#', true],
      [:logical_name, '論理カラム名', false, true, true],
      [:physical_name, '物理カラム名', false, true, true],
      [:type, '型', false, true, true],
      [:default, 'デフォルト値'],
      [:index, 'インデックス', true],
      [:option, 'オプション'],
      [:normalize, '正規化'],
      [:presence, '検証_必須', true],
      [:format, '検証_形式'],
    ]

    COLUMN_TYPES = %w(binary boolean date datetime decimal float integer string text time timestamp)

    RELATION_REV = {
      '1:1' => 'b',
      '1:n' => 'n:1',
      'n:1' => '1:n',
      'n:n' => 'n:n',
    }

    attr_reader :table

    def initialize(*files)
      if files.empty?
        die "error: .xls file is not found"
      end

      @table = {}
      @model_name = {}

      files.each do |file|
        puts "reading #{file}"
        load(file)
      end
    end

    private
    def load(file)
      relation_sheet = nil
      other_sheets = []
      Spreadsheet.open(file, 'rb') do |book|
        book.worksheets.each do |sheet|
          case sheet.name
          when 'テーブル一覧'
            load_table_list_sheet(sheet)
          when 'リレーション'
            relation_sheet = sheet
          else
            other_sheets << sheet 
          end
        end
      end

      load_relation_sheet(relation_sheet) if relation_sheet

      sheet_read = {}
      other_sheets.each do |sheet|
        if @model_name[sheet.name]
          load_table_sheet(sheet)
          sheet_read[sheet.name] = true
        else
          puts "ignoring sheet `#{sheet.name}'"
        end
      end

      @model_name.each_key do |sheet_name|
        unless sheet_read[sheet_name]
          die "error: sheet `#{sheet_name}' is not found"
        end
      end
    end

    def load_table_list_sheet(sheet)
      records = load_sheet(sheet, TABLE_LIST_FIELDS)

      records.each do |lineno, r|
        @table[r[:model_name]] = {
          :logical_name => r[:logical_name],
          :table_name => r[:table_name],
          :no_primary_key => r[:no_primary_key],
          :primary_key => r[:primary_key],
          :columns => [],
          :relations => [],
        }
        @model_name[r[:logical_name]] = r[:model_name]
      end
    end

    def load_table_sheet(sheet)
      records = load_sheet(sheet, COLUMN_FIELDS)

      records.each do |lineno, r|
        type = COLUMN_TYPES.find {|t| t =~ /\A#{r[:type]}/ }
        if type
          r[:type] = type
        else
          puts "warning: unknown type `#{r[:type]}' in sheet #{sheet.name}:#{lineno}"
        end

        # default value
        if !r[:default].empty?
          case r[:type]
          when 'integer', 'decimal'
            r[:default] = r[:default].to_i.to_s
          when 'float'
            r[:default] = r[:default].to_f.to_s
          end
        end
      end

      @table[@model_name[sheet.name]][:columns] += records.map {|lineno, r| r }
    end

    def load_relation_sheet(sheet)
      records = load_sheet(sheet, RELATION_FIELDS)

      records.each do |lineno, r|
        invalid_record = false
        [:table1, :table2, :through].each do |s|
          next if r[s].empty?
          unless @model_name[r[s]]
            puts "warning: ignore non-existent table `#{r[s]}' in sheet #{sheet.name}:#{lineno}"
            invalid_record = true
            next
          end
          r[s] = @model_name[r[s]]
          r[s] = r[s].pluralize if s == :through
        end
        next if invalid_record

        r[:relation].tr!('１一多ＭｍMmＮｎN対：', '11nnnnnnnn::')

        if r[:relation] == '1:1' && !r[:foreign1].empty?
          r[:relation] = 'b'
          relation_rev = '1:1'
        else
          relation_rev = RELATION_REV[r[:relation]]
        end

        unless relation_rev
          puts "warning: ignore invalid relation `#{r[:relation]}' in sheet #{sheet.name}:#{lineno}"
          next
        end

        foreign_key1 = nil
        if !r[:foreign1].empty? && r[:foreign1] != "1"
          foreign_key1 = r[:foreign1]
        end

        @table[r[:table1]][:relations] << {
          :table =>  r[:table2],
          :relation => r[:relation],
          :through => r[:through],
          :foreign => foreign_key1 || "",
          :option => r[:rel_option1],
        } 

        if !r[:foreign1].empty?
          @table[r[:table1]][:columns] << {
            :logical_name => "",
            :physical_name => foreign_key1 || r[:table2] + '_id',
            :type => 'integer',
            :default => r[:default1],
            :index => r[:index1],
            :option => r[:option1],
            :normalize => "",
            :presence => false,
            :format => "",
          }
        end

        foreign_key2 = nil
        if !r[:foreign2].empty? && r[:foreign2] != "1"
          foreign_key2 = r[:foreign2]
        end

        @table[r[:table2]][:relations] << {
          :table =>  r[:table1],
          :relation => relation_rev,
          :through => "",
          :foreign => foreign_key2 || "",
          :option => r[:rel_option2],
        } 

        if !r[:foreign2].empty?
          @table[r[:table2]][:columns] << {
            :logical_name => "",
            :physical_name => foreign_key2 || r[:table1] + '_id',
            :type => 'integer',
            :default => r[:default2],
            :index => r[:index2],
            :option => r[:option2],
            :normalize => "",
            :presence => false,
            :format => "",
          }
        end
      end
    end

    def load_sheet(sheet, fields)
      puts "reading sheet `#{sheet.name}'"
      header, *rows = active_rows(sheet)

      # read header
      field_colno = []
      header.each_with_index do |col, colno|
        i = fields.index {|f| f[1] == to_hankaku(col || "") }
        field_colno[i] = colno if i
      end

      fields.each_with_index do |field, i|
        if field[3] && field_colno[i].nil?
          die "error: header `#{field[1]}' is not found in sheet #{sheet.name}"
        end
      end

      # read rows
      records = []
      rows.each_with_index do |row, rowno|
        lineno = rowno + 2
        record = {}
        # read a row
        fields.each_with_index do |field, i|
          colno = field_colno[i]
          val = if !colno || !row[colno]
                  ""
                elsif row[colno].class != Float
                  row[colno].to_s.tr("¥", "\\")
                elsif row[colno] == row[colno].to_i
                  row[colno].to_i.to_s
                else
                  row[colno].to_f.to_s
                end
          if field[2] # boolean
            val = val !~ /\A\s*0?\s*\z/
          end
          record[field[0]] = val
          break if field[0] == :comment && val
          if field[4] && val.empty?
            die "erorr: value of `#{field[1]}' is empty in sheet #{sheet.name}:#{lineno}"
          end
        end

        records << [lineno, record] unless record[:comment]
      end

      records
    end

    def active_rows(sheet)
      rows = []
      sheet.each do |row|
        break if row.join =~ /\A\s*\z/
        rows << row 
      end
      rows
    end

    def die(mes)
      puts mes
      puts "aborted"
      exit 1
    end
  end

  module Migration
    class << self
      include Common

      def save(data_model)
        data_model.table.each do |name, table|
          columns = table[:columns]
          #next if columns.empty?

          file = filename(name)
          lines = IO.readlines(file)

          insert_column_lines(name, file, table, columns, lines)
          insert_index_lines(name, file, columns, lines)

          File.rename(file, "#{file}.bak")
          File.open(file, "w") {|f| f.puts lines }
          puts "modified migration file: #{file}"
        end
      end

      private
      def filename(name)
        file_pattern = "db/migrate/[0-9]*_create_#{name.pluralize}.rb"
        files = Dir.glob(file_pattern)
        if files.empty?
          puts "creating model files:"
          system "rails generate model #{name}"
          files = Dir.glob(file_pattern)
        elsif files.size > 1
          puts "warning: found more than one migration files #{file_pattern}"
        end

        files.last
      end

      def insert_column_lines(name, file, table, columns, lines)
        i_start = create_table_start_index(name, file, lines)
        replace_create_table_line(name, table, lines[i_start - 1]) if i_start

        i = index_and_delete(lines, regex(:CR))
        i ||= create_table_start_index(name, file, lines)

        lines.insert(i, *column_lines(columns)) if i
      end

      def replace_create_table_line(name, table, line)
        table_name = name.pluralize

        if table[:no_primary_key]
          s = ":id => false"
          line.sub!(/:id\s*=>\s*[^\s,]+/, s) ||
            line.sub!(/(create_table\s+:#{table_name})/, "\\1, #{s}")
        end

        if !table[:primary_key].empty?
          s = ":primary_key => :#{table[:primary_key]}"
          line.sub!(/:primary_key\s*=>\s*[^\s,]+/, s) ||
            line.sub!(/(create_table\s+:#{table_name})/, "\\1, #{s}")
        end
      end

      def create_table_start_index(name, file, lines)
        table_name = name.pluralize
        i = lines.index {|line| line =~ /\A\s*create_table\s+:#{table_name}[\s,]/ }
        if i
          i += 1
        else
          puts "#{file}: warning: `creata_table :#{table_name} ...' is not found"
        end
        i
      end

      def column_lines(columns)
        lines = []

        columns.each do |c|
          line = "      t.#{c[:type]} :#{c[:physical_name]}"
          if !c[:default].empty?
            line += ", :default => #{c[:default]}"
          end
          if !c[:option].empty?
            line += ", #{c[:option]}"
          end
          line += "  " + comment(:CR)
          lines << line 
        end

        lines
      end

      def insert_index_lines(name, file, columns, lines)
        table_name = name.pluralize
        i_change = index_and_delete(lines, regex(:CH))
        return unless columns.find {|c| c[:index] }

        unless i_change
          i_change = lines.index {|line| line =~ /\A\s*change_table\s+:#{table_name}\s/ }
          i_change += 1 if i_change
        end

        if i_change
          lines.insert(i_change, *index_lines(columns))
          return
        end

        i_create = create_table_start_index(name, file, lines)
        return unless i_create

        i_end = (i_create...lines.size).find {|i| lines[i] =~ /\A\s*end\s*\z/ }
        unless i_end
          puts "#{file}: warining: `end' is not found"
          return
        end
        index_lines = [
          "",
          "    change_table :#{table_name} do |t|",
          *index_lines(columns),
          "    end",
        ]
        lines.insert(i_end + 1, *index_lines)
      end

      def index_lines(columns)
        columns.select {|c| c[:index] }.map do |c|
          "      t.index :#{c[:physical_name]}  " + comment(:CH)
        end
      end
    end
  end

  module Model
    class << self
      include Common

      def save(data_model)
        data_model.table.each do |name, table|
          file = filename(name)
          lines = IO.readlines(file)

          columns = table[:columns]
          relations = table[:relations]
          insert_lines(name, file, regex(:NL), normalization_lines(columns), lines)
          insert_lines(name, file, regex(:VL), validation_lines(columns), lines)
          insert_lines(name, file, regex(:RL), relation_lines(relations), lines)
          insert_lines(name, file, regex(:DB), db_lines(table), lines)

          File.rename(file, "#{file}.bak")
          File.open(file, "w") {|f| f.puts lines }
          puts "modified model file: #{file}"
        end
      end

      private
      def filename(name)
        "app/models/#{name}.rb"
      end

      def insert_lines(name, file, comment_regex, lines_add, lines)
        i = index_and_delete(lines, comment_regex)
        lines_add << "" if !i && !lines_add.empty?
        i ||= class_start_index(name, file, lines)
        lines.insert(i, *lines_add) if i
      end

      def relation_lines(relations)
        lines = [] 
        return lines unless relations

        relations.each do |r|
          line = case r[:relation]
                 when '1:1'
                   "has_one :#{r[:table]}"
                 when 'b'
                   "belongs_to :#{r[:table]}"
                 when '1:n'
                   "has_many :#{r[:table].pluralize}" +
                     (!r[:through].empty? ? ", :through => :#{r[:through]}" : "")
                 when 'n:1'
                   "belongs_to :#{r[:table]}"
                 when 'n:n'
                   "has_and_belongs_to_many :#{r[:table].pluralize}"
                 else
                   raise "Program bug found!"
                 end
          if !r[:foreign].empty?
            line += ", :foreign_key => :#{r[:foreign]}"
          end
          if !r[:option].empty?
            line += ", #{r[:option]}"
          end
          lines << "  #{line}  #{comment(:RL)}"
        end

        lines
      end

      def validation_lines(columns)
        lines = []
        comment = "  #{comment(:VL)}"
        lines << "  # validation" + comment

        columns.each do |c|
          if c[:presence]
            lines << "  validates :#{c[:physical_name]}, :presence => true" + comment
          end
          if !c[:format].empty?
            lines << "  validates :#{c[:physical_name]}, :format => #{c[:format]}" + comment
          end
        end

        lines
      end

      def normalization_lines(columns)
        lines = []
        comment = "  #{comment(:NL)}"
        lines << "  # normalization" + comment

        columns.each do |c|
          next if c[:normalize].empty?
          name = c[:physical_name]
          block = [
            "def #{name}=(v)",
            "self[:#{name}] = v.#{c[:normalize]}",
            "end",
          ]
          lines << "  " + block.join('; ') + "#{comment}"
        end

        lines
      end

      def db_lines(table)
        lines = []
        if !table[:table_name].empty?
          lines << "  set_table_name :#{table[:table_name]}  #{comment(:DB)}"
        end
        if !table[:primary_key].empty?
          lines << "  set_primary_key :#{table[:primary_key]}  #{comment(:DB)}"
        end
        lines
      end

      def class_start_index(name, file, lines)
        class_name = name.camelize
        i = lines.index {|line| line =~ /\A\s*class\s+#{class_name}\s/i }
        if i
          i += 1
        else
          puts "#{file}: warning: `class #{class_name} ...' is not found"
        end
        i
      end
    end
  end

  module Locale
    class << self
      include Common

      def save(data_model)
        file = filename
        lines = IO.readlines(file)
        insert_model_lines(name, file, data_model.table, lines)

        data_model.table.reverse_each do |name, table|
          insert_attr_lines(name, file, table[:columns], lines)
        end

        replace_file(file, lines)
        puts "modified locale file: #{file}"
      end

      private
      def filename
        file = "config/locales/translation_ja.yml"
        unless File.exist?(file)
          File.open(file, "w") do |f|
            f.puts "ja:"
            f.puts "  activerecord:"
            f.puts "    models:"
            f.puts ""
            f.puts "    attributes:"
          end
        end
        file
      end

      def insert_model_lines(name, file, table, lines)
        i = index_and_delete(lines, regex(:LCM))
        i ||= model_start_index(file, lines)

        lines.insert(i, *model_lines(table)) if i
      end

      def model_lines(table)
        table.map do |name, t|
          "      #{name}: \"#{t[:logical_name]}\"  #{comment(:LCM)}"
        end
      end

      def model_start_index(file, lines)
        i = lines.index {|line| line == "    models:\n" }
        if i
          i += 1
        else
          puts "#{file}: warning: `    models:' is not found"
        end
        i
      end

      def insert_attr_lines(name, file, columns, lines)
        i = index_and_delete(lines, /#{comment(:LCA, name)}\s/)
        i ||= attr_start_index(name, file, lines)

        lines.insert(i, *attr_lines(name, columns)) if i
      end

      def attr_lines(name, columns)
        columns.select {|c| !c[:logical_name].empty? }.map do |c|
          "        #{c[:physical_name]}: \"#{c[:logical_name]}\"  #{comment(:LCA, name)}"
        end
      end

      def attr_start_index(name, file, lines)
        i_attr = lines.index {|line| line == "    attributes:\n" }
        unless i_attr
          puts "#{file}: warning: `    attributes:' is not found"
          return nil
        end
        i_attr += 1

        i_model = (i_attr...lines.size).find {|i| lines[i] =~ /\A      #{name}:/}
        if i_model
          i_model += 1
        else
          lines.insert(i_attr, "      #{name}:")
          i_model = i_attr + 1
        end
        i_model
      end
    end
  end
end

task :tablis do
  data_model = Tablis.load(*Dir.glob('config/tablis_*.xls'))
  Tablis::Migration.save(data_model)
  Tablis::Model.save(data_model)
  Tablis::Locale.save(data_model)

  # TODO: Locale: ja.yml or translation_ja.yml?
  # TODO: modify test Excel file
  # TODO: create gem
end
