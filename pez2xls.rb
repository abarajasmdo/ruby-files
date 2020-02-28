def intro()
puts "       *****************************************************************"
puts " "
puts "       PLEASE DO NOT CLOSE THIS WINDOW OR TRIED TO OPEN A EXCEL DOCUMENT"
puts " "
puts "       *****************************************************************"
puts " "
puts "       ---------------------    PLEASE WAIT!!!     ---------------------"
puts " "
puts "       *****************************************************************"
puts " "
puts " "
puts "       #################################################################"
puts " "
puts "            THE PROGRAM IS CONVERTING THE *.PEZ IN *.XLS (FILE PER FILE)"
puts "            THE PROGRAM SEPARATES PER FILE ONE SHEET PER SURFSEAL REGION"
puts " "
puts "       #################################################################"
puts " "
puts " "
puts " "
puts "       _________________________________________________________________"
puts "                                    CREATED BY ANTONIO BARAJAS - NHCUKHT"
puts "                                                           NOVEMBER 2007"
end
module EC
end
def array2excel(b)
  require 'win32ole'            	                                            #SE ABRE APLICACIÓN DE EXCEL
  excel = WIN32OLE.new("excel.application")
  WIN32OLE.const_load(excel, EC)                                      	#CARGA CONSTANTES LOCALES DE EXCEL A RUBY
  excel['Visible'] = FALSE;                                                #HACE VISIBLE EL EXCEL Y CREA UN LIBRO DE EXCEL
  b.keys.sort.each do |book|
    workbook = excel.Workbooks.Add(EC::XlWBATWorksheet);
    b[book].keys.sort.reverse.each do |sht|
      sheet = workbook.Worksheets.Add();
      sheet.Name = sht;
      n=1
      col=Array.new;
      col=[" ","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p"]      
      b[book][sht].each do |line|
        m=1
        case n
          when 1...5 then
            yy=col[m];
            excel.Range("#{yy}#{n}")['Value'] = line;
          when 5 then
            #~ lreg=Array.new
            #~ reg_exp=/(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(.*)(?>\s*)?/
            #~ if line =~ reg_exp then
              #~ lreg<<$1<<$2<<$3<<$4<<$5<<$6<<$7<<$8<<$9<<$10<<$11<<$12<<$13
              #~ lreg.each do |data|
                #~ yy=col[m];
                #~ excel.Range("#{yy}#{n}")['Value'] = data;
                #~ m += 1
              #~ end
            #~ end
          when 6 then
            #~ lreg=Array.new
            #~ reg_exp=/(?>\s*)?(\d*)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(\d*)(?>\s*)?(\d*)(?>\s*)?(.*)(\(.*\))(?>\s*)?(.*)/
            #~ if line =~ reg_exp then
              #~ lreg<<$1<<$2<<$3<<$4<<$5<<$6<<$7<<$8<<$9<<$10<<$11<<$12<<$13
              #~ lreg.each do |data|
                #~ yy=col[m];
                #~ excel.Range("#{yy}#{n}")['Value'] = data;
                #~ m += 1
              #~ end
            #~ end
          when 7 then
            lreg=Array.new
            reg_exp=/(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(.*)(?>\s*)?/
            if line =~ reg_exp then
              lreg<<$1<<$2<<$3<<$4<<$5<<$6<<$7<<$8<<$9<<$10<<$11<<$12<<$13
              lreg.each do |data|
                yy=col[m];
                excel.Range("#{yy}#{n}")['Value'] = data;
                m += 1
              end
            end
          else
            lreg=Array.new
            reg_exp=/(?>\s*)?(\d*)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(\d*)(?>\s*)?(\d*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)/
            if line =~ reg_exp then
              lreg<<$1<<$2<<$3<<$4<<$5<<$6<<$7<<$8<<$9<<$10<<($11+$12+$13)
              lreg.each do |data|
                yy=col[m];
                excel.Range("#{yy}#{n}")['Value'] = data;
                m += 1
              end
            end
        end
        n += 1
      end
    end
    workbook.Sheets('Sheet1').Select;
    main = workbook.ActiveSheet;
    main.Name = "RESUME";
    n=1
    col=Array.new;
    col=[" ","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p"]      
    b[book].keys.sort.each do |sht|
      m=1
      lreg=Array.new
      if n==1 then
        title=b[book][sht][4]
        reg_exp=/(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(.*)(?>\s*)?/
        if title =~ reg_exp then
          lreg<<$1<<$2<<$3<<$4<<$5<<$6<<$7<<$8<<$9<<$10<<$11<<$12<<$13
          lreg.each do |data|
            yy=col[m];
            excel.Range("#{yy}#{n}")['Value'] = data;
            m += 1
          end
        end
        n += 1
      end
      m=1
      line=b[book][sht][5]
      yy=col[m];
      excel.Range("#{yy}#{n}")['Value'] = sht;
      m += 1
      lreg=Array.new
      reg_exp=/(?>\s*)?(\d*)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(-?\d*\.\d)(?>\s*)?(\d*)(?>\s*)?(\d*)(?>\s*)?(?>MAX ON )(.*)(\(.*\))(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)/
      if line =~ reg_exp then
        lreg<<$1<<$2<<$3<<$4<<$5<<$6<<$7<<$8<<$9<<$10<<$11<<$12<<($13+$14+$15)
        lreg.each do |data|
          yy=col[m];
          excel.Range("#{yy}#{n}")['Value'] = data;
          m += 1
        end
      end
      n += 1
    end
    workbook.SaveAs(Fld+"/"+book+".xls")
    excel.ActiveWorkbook.Close(0);
  end
  excel.Quit();
end
def get_pezdata()
  b=Hash.new
  pezfiles = `ls *.pez`.split(".pez\n")
  pezfiles.each do |pez|
    a=File.read(pez+".pez").split("####################################################################################################################################\n")
    b[pez]=Hash.new  
    n=0
    a.each do |patch|
      if n>0 then
        pnam=patch.split("\n")[0].to_s
        b[pez][pnam]=Array.new
        b[pez][pnam]=patch.split("\n")
        b[pez][pnam].delete(pnam)
        b[pez][pnam][5]=b[pez][pnam][4]
        b[pez][pnam][4]="ELEM SIGX SIGY TXY SIG1 SIG2 SIGE TEMP COL ROW CRITERIA COORD LOCATION XDIR:-"
        b[pez][pnam].delete("")
      end
      n=1
    end
  end
  return b
end

intro()
raiz=`pwd`.split("\n")[0]       #OBTENER LA RAIZ DE LA CARPETA DE TRABAJO
drive=`vol`.split[3]                #OBTENER EL DISCO DE TRABAJO
Fld=drive+":"+raiz                #GENERAR LA DIRECCION DE TRABAJO
b=get_pezdata()
array2excel(b)