require 'axlsx'

module ExcelTable
  class Worksheet
    attr_accessor :rows, :headings
    attr_reader :save_filename

    def initialize(params={})
      @package = Axlsx::Package.new
      @workbook = @package.workbook
      prepare_styles
      @headings = params.fetch(:headings, [])
      @rows = params.fetch(:rows, [])
    end

    def save(params={})
      style = params.fetch(:style, false)
      name = params.fetch(:name){ | key | warn "sheet #{key} not provided, use `Sheet1' by default"; 'Sheet1'}
      params.merge!(:name => name)
      @workbook.add_worksheet(params) do |sheet|
        # Applies the black_cell style to the first and third cell, and the blue_cell style to the second.
        sheet.add_row headings, :style => @header_row unless headings.empty?

        # Applies the thin border to all three cells
        rows.each_index do |index|
          sheet.add_row rows[index], :types => :string
        end

        add_style(sheet, headings) if style
      end

      @save_filename = params.fetch(:save_filename, 'default.xlsx')
      @package.serialize(@save_filename)
    end

    def to_str
      @package.to_stream.read
    end

    private

    def prepare_styles
      @workbook.styles do |s|
        left_border = {border: {edges: [:left, :bottom], style: :thin, :color => '00'}}
        right_border = {border: {edges: [:right, :bottom], style: :thin, :color => '00'}}

        @odd_first_column_style = s.add_style odd_row_style.merge(left_border)
        @odd_last_column_style =s.add_style odd_row_style.merge(right_border)

        @even_first_column_style = s.add_style even_row_style.merge(left_border)
        @even_last_column_style =s.add_style even_row_style.merge(right_border)


        @header_row = s.add_style :bg_color => '00', :fg_color => 'FF', :sz => 12, :alignment => {:horizontal => :left}
        @odd_row = s.add_style odd_row_style
        @even_row = s.add_style even_row_style
      end
    end

    def add_style(sheet, headings)
      first_row = headings.empty? ? 0 : 1
      (first_row..sheet.rows.size - 1).each do |index|
        row = sheet.rows[index]
        if index.odd?
          row.style = @even_row
          row.cells.first.style = @even_first_column_style
          row.cells.last.style = @even_last_column_style
        else
          row.style = @odd_row
          row.cells.first.style = @odd_first_column_style
          row.cells.last.style = @odd_last_column_style
        end
      end
    end

    def odd_row_style
      {
          bg_color: 'ff',
          fg_color: '00',
          sz: 12,
          alignment:
              {horizontal: :left},
          border:
              {edges: [:bottom],
               style: :thin,
               color: '00'}
      }
    end

    def even_row_style
      {
          bg_color: 'CDCDCD',
          fg_color: '00',
          sz: 12,
          alignment:
              {horizontal: :left},
          border:
              {edges: [:bottom],
               style: :thin,
               color: '00'}
      }
    end

  end
end