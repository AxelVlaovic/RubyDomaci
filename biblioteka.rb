require 'roo'
require 'roo-xls'

class Table

  include Enumerable

  def initialize (*args)


    if args[0].is_a? Array then

      @table = []
      @table = args[0]
      @columns = []

      @columns = columns_generate 

      (0..@table[0].length-1).each do |index|
          Table.create_method(@table[0][index])
      end

    else
      
    path = args[0]
      
    excel = Roo::Spreadsheet.open(path,{:expand_merged_ranges => true})

    @table = []
    @columns = []

    excel.sheets.each do |sheet_name|
      sheet = excel.sheet(sheet_name)
      if !sheet.nil?
            if !sheet.last_row.nil? && !sheet.last_column.nil?

                left = first_data_col_index sheet.row(sheet.first_row)
                right = last_data_col_index sheet.row(sheet.first_row)

                (sheet.first_row..sheet.last_row).each do |i|

                  if(valid sheet.row(i)) then

                    table_row = data_row sheet.row(i),left,right
                    @table << table_row

                  end
                  
                end

            end
      end

    end
   
    @columns = columns_generate 

    (0..@table[0].length-1).each do |index|
        Table.create_method(@table[0][index])
    end

  end

  end


  def self.create_method(name)
    define_method("#{name}") do 
      
      result_column = []

      for i in 0..@columns.length-1 do
      
        if(@columns[i][0] == name) then

          for j in 1..@columns[i].length-1 do
              result_column << @columns[i][j]            
          end

        end

      end
            #Ne radi
      for i in 0..@columns.length-1 do
      
        if ((@columns[i][0] == name) && (@columns[i].all?(String))) then
          
          for l in 0..result_column.length-1 do
            for z in 0..@table.length-1 do
              
              if(@table[z].include? result_column[l])
                
                arr = Array.new
                arr = @table[z]
                id = result_column[l]

                result_column.define_singleton_method(id.to_sym) {arr}

              end

            end
          end

        end
      end

      return result_column

      return nil

    end

    define_method("#{name}.sum") do 
      
      result_column = []

      for i in 0..@columns.length-1 do
      
        if(@columns[i][0] == name) then

          for j in 1..@columns[i].length-1 do
              result_column << @columns[i][j]            
          end

        end

      end

      sum = 0
      flag = false

      result_column.each do |element|

        if(element.is_a? Integer)
          sum+=element
          flag = true
        end

      end

      if(flag == true)
        return sum
      end

      return nil

    end
    
  end


  def [] header

    test_header = ""

    if (header.is_a? Symbol)
      test_header = header.to_s
    else
      test_header = header
    end
  
    result_column = []

      for i in 0..@columns.length-1 do
      
        if(@columns[i][0] == test_header) then

          for j in 1..@columns[i].length-1 do
              result_column << @columns[i][j]            
          end

          return result_column 

        end

      end

      return nil

  end

  def valid row

    not_an_empty_row = false

    row.each do |element|

      if (!element.nil?)
          not_an_empty_row = true
          break
      end

    end

    row.each do |element|

      if (element.is_a? String)
        el = element.downcase
        if((el == "total") || (el == "subtotal"))
          return false
        end
      end

    end

    return not_an_empty_row

  end

  def first_data_col_index headers

    for i in 0..headers.size-1 do

      if(!headers[i].nil?)
          return i
      end

    end

  end

  def last_data_col_index headers

    for i  in headers.size.downto(0) do
      
        if(!headers[i].nil?)
            return i
        end

    end

  end

  def data_row row,left,right 

    arr = []

    (left..right).each do |i|
      arr << row[i]
    end

    return arr

  end

  def columns_generate  

    columns = []
    columnsTemp = []
    number_of_columns = @table[0].length

    for  n in 0..number_of_columns-1 do 
      for i in 0..@table.length-1 do
          columnsTemp << @table[i][n]
      end

      columns << columnsTemp
      columnsTemp = []
    end

    return columns

  end

  def each 
    for i in 0..@table.length-1 do
      for j in 0..@table[i].length-1 do
            yield @table[i][j]        
      end
    end
  end

  def row index
    @table[index]
  end

  def table 
    @table
	end

  def + (table2)

    if(@table[0] == table2.table[0]) then
    
      for i in 1..table2.table.length-1 do
        
        @table << table2.table[i]

      end
    else
      return "Ove tabele nemaju iste headere"
    end


    return @table

    
  end

  def - (table2)

    if(@table[0] == table2.table[0]) then
      for i in 1..table2.table.length-1 do
        
        for j in 1..@table.length-1 do

          if(table2.table[i] == @table[j])
            @table.delete(table2.table[i])
          end

        end

      end
    else
      return "Ove tabele nemaju iste headere"
    end

    return @table

  end

end





