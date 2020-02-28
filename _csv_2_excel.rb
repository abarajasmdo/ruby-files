module EC
end
def intro()
puts "               ********** WARNING *************

 #####################################################################
   RRRRRR   PPPPPP   EEEEEE        GGGGGG   EEEEEE   IIIIII   QQQQQQ
   RR  RR   PP  PP   EE            GG       EE         II     QQ  QQ
   RRRRRR   PPPPPP   EEEEE         GGGGGG   EEEEE      II     QQ  QQ
   RR RR    PP       EE            GG  GG   EE         II     QQ  QQ
   RR  RR   PP       EEEEEE        GGGGGG   EEEEEE   IIIIII   QQQQQQQ
 Rotating Pars Engineering at General Electric Infrastructure Queretaro
 #####################################################################

  This is Version I.0 to compress csv files to excel sheet.
  If you encounter an error, please call me.
       Antonio Barajas (8*879-0611).
       
  Program build on January 2011
"
end
def getfolder()
  folder=`cd`.split("\n")[0]#.gsub("\\","/")       #OBTENER LA RAIZ DE LA CARPETA DE TRABAJO
  return folder
end
def csv2excel()
  excel = WIN32OLE.new("excel.application")
  excel.DisplayAlerts = false
  excel.Interactive = false
  excel.ScreenUpdating = false
  excel['Visible'] = false
  WIN32OLE.const_load(excel, EC)
  b=`ls *.csv`.split("\n")
  workbook = excel.Workbooks.Add(EC::XlWBATWorksheet)
  i=0
  b.each do |sheet|
    if i>0 then
      sht = workbook.Worksheets.Add()
      sht.Name = sheet.gsub(".csv","").split("loc")[1].to_s
    else
      workbook.Sheets('Sheet1').Select
      sht = workbook.ActiveSheet
      sht.Name = sheet.gsub(".csv","").split("loc")[1].to_s
    end      
    wtemporal=excel.workbooks.open(Fld+"/"+sheet)
    wtemporal.ActiveSheet.Range("A1:Z200").Copy
    sht.Paste
    wtemporal.Close(0)
    i=1
  end
  workbook.SaveAs(Fld+"/"+b[0].gsub(".csv","").split("_loc")[0]+".xls")
  excel.Interactive = true
  excel.ScreenUpdating = true
  excel.ActiveWorkbook.Close(0)
  excel.Quit();
end
#~ require 'hpricot'
#~ require 'watir'
require 'win32ole'
intro()
Fld=getfolder()
csv2excel()
