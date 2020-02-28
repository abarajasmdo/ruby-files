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

  This is Version I.0 of Get Life Comparison for TPS Process.
  If you encounter a real error, please call me.
       Antonio Barajas (8*879-0611).
       
  Program build on March 2010
"
end
def getfolder()
  folder=`cd`.split("\n")[0].gsub("\\","/")       #OBTENER LA RAIZ DE LA CARPETA DE TRABAJO
  return folder
end
def csv2excel(labl,tmscr)
  excel = WIN32OLE.new("excel.application")
  excel.DisplayAlerts = false
  excel.Interactive = false
  excel.ScreenUpdating = false
  excel['Visible'] = false
  if tmscr==0 then
    WIN32OLE.const_load(excel, EC)
  end
  b=`ls *.csv`.split("\n")
  workbook = excel.Workbooks.Add(EC::XlWBATWorksheet)
  i=0
  b.each do |sheet|
    if i>0 then
      sht = workbook.Worksheets.Add()
      sht.Name = sheet.gsub(".csv","").split("_")[1].to_s
    else
      workbook.Sheets('Sheet1').Select
      sht = workbook.ActiveSheet
      sht.Name = sheet.gsub(".csv","").split("_")[1].to_s
    end      
    wtemporal=excel.workbooks.open(Fld+"/"+sheet)
    wtemporal.ActiveSheet.Range("A1:Z100").Copy
    sht.Paste
    wtemporal.Close(0)
    i=1
  end
  workbook.SaveAs(Fld+"/"+labl+".xls")
  excel.Interactive = true
  excel.ScreenUpdating = true
  excel.ActiveWorkbook.Close(0)
  excel.Quit();
end
#~ require 'hpricot'
require 'win32ole'
require 'watir'
intro()
Fld=getfolder()
csv2excel("compress",0)
