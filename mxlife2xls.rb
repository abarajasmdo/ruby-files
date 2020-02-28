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
puts "         THE PROGRAM IS CONVERTING THE *.mxlife IN *.XLS (FILE PER FILE)"
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
          when 1 then
            yy=col[m];
            excel.Range("#{yy}#{n}")['Value'] = line.strip;
          when 2 then
            lreg=Array.new
            reg_exp=/(?>\s*)?(THE PATCH MAX SEFF RANGE IS )(?>\s*)?(\d*.\d)(?>\s*)?(KSI)(?>\s*)?(AT: EL2D)(?>\s*)?(\d*)(?>\s*)?(\w*)/
            if line =~ reg_exp then
              lreg<<$1<<$2<<$3<<$4<<$5<<$6
              lreg.each do |data|
                yy=col[m];
                excel.Range("#{yy}#{n}")['Value'] = data;
                m += 1
              end
            end
          when 3 then
            lreg=Array.new
            reg_exp=/(?>\s*)?(3D ELEM )(?>\s*)?(\d*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?/
            if line =~ reg_exp then
              lreg<<$1<<$2<<$3<<4
              lreg.each do |data|
                yy=col[m];
                excel.Range("#{yy}#{n}")['Value'] = data;
                m += 1
              end
            end
          when 4 then
            lreg=Array.new
            reg_exp=/(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)(?>\s*)?(\w*)/
            if line =~ reg_exp then
              lreg<<$1<<$2<<$3<<$4<<$5<<$6<<$7
              lreg.each do |data|
                yy=col[(m+1)];
                excel.Range("#{yy}#{n}")['Value'] = data;
                m += 1
              end
            end
          else
            lreg=Array.new
            lreg2=line.split(" ")[0]
            reg_exp=/(?>\.*\s*)?(\d*.\d)(?>\s*)?(-?\d*.\d)(?>\s*)?(-?\d*.\d)(?>\s*)?(-?\d*.\d)(?>\s*)?(-?\d*.\d)(?>\s*)?(-?\d*.\d)(?>\s*)?(-?\d*.\d)(?>\s*)?/
            if line =~ reg_exp then
              lreg<<$1<<$2<<$3<<$4<<$5<<$6<<$7
              if lreg2==lreg[0] then
              else
                yy=col[m];
                excel.Range("#{yy}#{n}")['Value'] = lreg2;
              end
              lreg.each do |data|
                yy=col[m+1];
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
      yy=col[m];
      excel.Range("#{yy}#{n}")['Value'] = sht;
      n += 1
      m += 1
      line=b[book][sht][1]
      lreg=Array.new
      reg_exp=/(?>\s*)?(THE PATCH MAX SEFF RANGE IS )(?>\s*)?(\d*.\d)(?>\s*)?(KSI)(?>\s*)?(AT: EL2D)(?>\s*)?(\d*)(?>\s*)?(\w*)/
      if line =~ reg_exp then
        lreg<<$1<<$2<<$3<<$4<<$5<<$6
        lreg.each do |data|
          yy=col[m];
          excel.Range("#{yy}#{n}")['Value'] = data;
          m += 1
        end
        n+=1
      end

      line=b[book][sht][2]
      m=2
      lreg=Array.new
      reg_exp=/(?>\s*)?(3D ELEM )(?>\s*)?(\d*)(?>\s*)?(.*)(?>\s*)?/
      if line =~ reg_exp then
        lreg<<$1<<$2<<$3
        lreg.each do |data|
          yy=col[m];
          excel.Range("#{yy}#{n}")['Value'] = data;
          m += 1
        end
        n += 1
      end
    end
    workbook.SaveAs(Fld+"/"+book+".xls")
    excel.ActiveWorkbook.Close(0);
  end
  excel.Quit();
end
def get_pezdata()
  b=Hash.new
  pezfiles = `ls f60.mxlife`.split(".mxlife\n")
  pezfiles.each do |pez|
    a=File.read(pez+".mxlife").split("\n     SURFSEAL")
    b[pez]=Hash.new  
    n=0
    a.each do |patch|
      if n>0 then
        pnam2=patch.split("\n")[0]
        pnam=patch.split("\n")[0].to_s.strip.sub(/(PATCH      )/, 'Region')
        b[pez][pnam]=Array.new
        b[pez][pnam]=patch.split("\n")
        b[pez][pnam].delete(pnam2)
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
