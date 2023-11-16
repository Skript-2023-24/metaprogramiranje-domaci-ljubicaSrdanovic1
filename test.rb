require 'google_drive'

class GoogleSheets
  include Enumerable

  def initialize(worksheet)
    @worksheet = worksheet
  end

  def each(&block)
    table_values.each(&block)
  end

  def row(index)
    table_values[index]
  end

  def num_rows
    table_values.length
  end

  def [](column_name)
    column_data = table_values.transpose[header_row.index { |header| header.downcase == column_name.downcase }]
    column_data || []
  end

  def prvaKolona
    ClassColumn.new(self, 'Prva Kolona')
  end

  def drugaKolona
    ClassColumn.new(self, 'Druga Kolona')
  end

  def trecaKolona
    ClassColumn.new(self, 'Treca Kolona')
  end

  def self.new_from_values(table_values, header_row)
    new_instance = allocate
    new_instance.instance_variable_set(:@table_values, table_values)
    new_instance.instance_variable_set(:@header_row, header_row)
    new_instance
  end

  class ClassColumn
    def initialize(helper, column_name)
      @helper = helper
      @column_name = column_name
    end

    def map(&block)
      @helper[@column_name].map { |cell| cell.empty? ? cell : block.call(cell) }
    end

    def select(&block)
      @helper[@column_name].select(&block)
    end

    def reduce(initial_value, &block)
      @helper[@column_name].reduce(initial_value, &block)
    end

    def join(separator = ', ')
      @helper[@column_name].join(separator)
    end
  end

  def sum(column_name)
    column = self[column_name].map(&:to_f).compact
    column.sum
  end

  def avg(column_name)
    column = self[column_name].map(&:to_f).reject(&:zero?)
    column.sum / column.length
  end

  def rn(index)
    identifier_column_data = table_values.transpose[header_row.index { |header| header.downcase == 'indeks' }]
    target_row_index = identifier_column_data&.index("rn#{index}")
    target_row_index ? table_values[target_row_index] : nil
  end

  protected

  def table_values
    @table_values ||= begin
    start_row, start_col = find_first_cell_with_value(@worksheet)
    return [] unless start_row

    @header_row = (start_col..@worksheet.num_cols).map { |col_index| @worksheet[start_row, col_index].to_s }

    (start_row..@worksheet.num_rows).map do |row_index|
      row_values = (start_col..@worksheet.num_cols).map { |col_index| @worksheet[row_index, col_index].to_s }
      row_values if !row_values.any? { |cell| cell.downcase.include?('total') || cell.downcase.include?('subtotal') }
    end.compact
    end
  end

  def header_row
    @header_row ||= []
  end

  def find_first_cell_with_value(worksheet)
    (1..worksheet.num_rows).each do |row_index|
      (1..worksheet.num_cols).each do |col_index|
        return [row_index, col_index] if worksheet[row_index, col_index] && !worksheet[row_index, col_index].empty?
      end
    end

    [nil, nil]
  end
end

def main
  session = GoogleDrive::Session.from_config('config.json')
  ws = session.spreadsheet_by_key('1Q984joayy3P2PegoCLQS3dCWGAvGRq4aFX2QKuLbSFQ').worksheets[0]

  t = GoogleSheets.new(ws)

  ws_new = session.spreadsheet_by_key('1Q984joayy3P2PegoCLQS3dCWGAvGRq4aFX2QKuLbSFQ').worksheets[1]
  t_new = GoogleSheets.new(ws_new)


  puts "Dvodimenzionalni niz sa vrednostima tabele:"
  t.each { |row| puts "[#{row.join(', ')}]" }
  puts "\n"
  
  #puts "\nPristup redovima preko t.row(indeks):"
  #p t.row(2)

  #puts "\nSve ćelije u tabeli:"
  #t.each { |row| puts "#{row.join(', ')}" }

  #puts "\nPristup koloni(t[“Prva Kolona”]):"
  #p t["Prva kolona"]

  #puts "\nPristup vrednostima unutar kolone (t['Prva kolona'][3]):"
  #p t["Prva kolona"][3]

  #puts "\nPromena unutar kolone (t['Prva kolona'][3]):"
  #t["Prva kolona"][3]= 2556
  #p t["Prva kolona"][3].to_s

  #p "Prva Kolona: #{t.prvaKolona.join(', ')}"

  #p "Suma Prva kolona: #{t.sum('Prva kolona')}"
  #p "Avg Prva kolona: #{t.avg('Prva kolona')}"
  
  #puts "\nIzvlačenje reda na osnovu indeksa (t.rn2310):"
  #result_row = t.rn('2310')

  #if result_row
  #  puts "Red za indeks '2310': #{result_row.join(', ')}"
  #else
  #  puts "Red za indeks '2310' nije pronađen."
  #end

  #mapped_values = t.prvaKolona.map { |cell| cell.to_i + 1 }
  #puts "Mapirane vrednosti: #{mapped_values.join(', ')}"

  #selected_values = t.prvaKolona.select { |cell| cell.to_i > 5 }
  #puts "Selektovane vrednosti: #{selected_values.join(', ')}"

  #reduced_value = t.prvaKolona.reduce(1) { |sum, cell| sum + cell.to_i }
  #puts "Reduce vrednost: #{reduced_value}"


end

puts "\n"

main