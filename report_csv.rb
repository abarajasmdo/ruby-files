def transponse_array(entry)
  # Transponse Table
  array=Array.new
  mm=0
  entry.each do |c|
      m=0
      c.each do |w|
        if mm==0 then
          array[m]=Array.new
        end
        array[m]<<w
        m+=1
      end
      mm=1
    end
  return array
end
def htmltable_toarray(file="C:\Documents and Settings\nhcukht\Desktop\report.xls")
  require 'hpricot'
  require 'open-uri'
  # load the report.xls generated for Orion
  doc = Hpricot(open(file)).at("table")
  # Create all the required Arrays
  a=Array.new
  b=Array.new
  matrix=Array.new
  filas=Array.new
  # Find all the headers of the Table and put it into a matrix
  a= doc.search("th").to_a
  a.each do |x|
      matrix<<x.to_plain_text.gsub("\n","   ").gsub("?"," ").gsub(",",";")
    end
  filas<<matrix
  # Find all the Rows of the Table
  n=1
  b= doc.search("tr").to_a
  b.each do |y|
      filas[n]=Array.new
      y.search("td").to_a.each do |z|
        filas[n]<<z.to_plain_text.gsub("\n","   ").gsub("?"," ").gsub(",",";")
      end
      n+=1
    end
    return filas
end
def array_tocsv(array,filename)
  filename=filename+".csv"
  afile = File.new(filename,"w")
  x=Array.new
  array.each do |x|
      x.each do |y|
        afile<<y+","
      end
      afile<<"\n"
    end
end

table=htmltable_toarray()
array=transponse_array(table)
array_tocsv(array,"C:\Documents and Settings\nhcukht\Desktop\report_t")