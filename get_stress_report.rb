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

  This is Version I.0 of Extract Stress for TPS Process.
  If you encounter a real error, please call me.
       Antonio Barajas (8*879-0611).
       
  Program build on February 2009
"
end
def getfolder()
  folder=`pwd`.split("\n")[0]       #OBTENER LA RAIZ DE LA CARPETA DE TRABAJO
  if folder.split("/")[1].size > 1 then
      drive=`vol`.split[3] + ":"
      folder=drive+folder
  else
    if folder=~/^\/(\w+)\// then
      drive=$1+":/"
    end
    folder[0,3]=drive
  end
  return folder
end
def inputfiles()
  h=File.new("mcas.lock","w+")
  h.close
   if FileTest.exist?("tables.dat") == true
      a=File.read("tables.dat").squeeze(" ").split("TABL")
      b=Array.new
      i=0
      a.each do |data|
        if data==" " then
        else
          b[i]=Array.new
          b[i]<<data.split[0]<<data.gsub(".","").split[1]<<data.gsub("\n ","").split("FOR")[1]
          i=i+1
        end
      end
  end
  c=`ls *.37`.split[0].to_s
  p c
  g=File.new("report_run.bat","w+")
  f=File.new("table_list.csv","w+")
  b.each do |line|
    f<<line[0]<<","<<line[1]<<","<<line[2]<<"\n"
    if line[1]=="CENT" then
      line[1]="13 14"
    elsif line[1]=="SURFACE" then
      line[1]="19 20"
    end
    g<<"siesta_lite << log
mcas
"+c+"

6
l
h
"+line[0]+"    
"+line[1]+"


fini
log
mv f56.dat loc_"<< line[0] <<".dat
rm f37.dat
rm f45.dat
"
  end
  g<<"rm mcas.lock
"
  f.close
  g.close
end
def getinformation()
  files=`ls loc_*.dat`.split
  files.each do |f|
    ff=f.gsub("loc_","").gsub(".dat","").to_i
    if ff<10 then
      fff="00"+f.gsub("loc_","").gsub(".dat","")
    elsif ff<100 then
      fff="0"+f.gsub("loc_","").gsub(".dat","")
    elsif ff<1000 then
      fff=f.gsub("loc_","").gsub(".dat","")
    end
    i=1
    File.read(f).gsub("\n\n\n","\n $\n $\n").gsub(/^ \$(\s.*)(\n)/,"").gsub("         ","    NA    ").gsub(/^      TIME(\s.*)(\n)/,"$\n").gsub(/^      MAX(\s.*)(\n)/,"$\n").gsub(/^      MIN(\s.*)(\n)/,"$\n").gsub("\n\n","\n").gsub("CENT ","CENT_").squeeze(" ").gsub("\n ","\n").gsub(/^ /,"").gsub("LOC ","LOC_").gsub("FACE ","FACE_").gsub(" \n","\n").gsub(" ",",").gsub("CENT_,","CENT,").gsub("CENT_\n","CENT\n").gsub("NA","").split("$\n$\n$\n$\n").each do|loc|
      if i==1 then
        h=File.new("ss_"+fff+".csv","w+")
        h<<loc
        h.close
      else
        h=File.new("ss_"+fff+"_"+i.to_s+".csv","w+")
        h<<loc
        h.close        
      end
      i=i+1
    end
  end
end
def csv2excel()
  excel = WIN32OLE.new("excel.application")
  excel.DisplayAlerts = false
  excel.Interactive = false
  excel.ScreenUpdating = false
  excel['Visible'] = false
  WIN32OLE.const_load(excel, EC)
  b=`ls ss_*.csv`.split("\n")
  workbook = excel.Workbooks.Add(EC::XlWBATWorksheet)
  i=0
  b.reverse_each do |sheet|
    if i>0 then
      sht = workbook.Worksheets.Add()
      sht.Name = sheet.gsub(".csv","")
    else
      workbook.Sheets('Sheet1').Select
      sht = workbook.ActiveSheet
      sht.Name = sheet.gsub(".csv","")
    end      
    wtemporal=excel.workbooks.open(Fld+"/"+sheet)
    wtemporal.ActiveSheet.Range("A1:AZ6000").Copy
    sht.Paste
    wtemporal.Close(0)
    i=1
  end
  workbook.SaveAs(Fld+"/ss_summary.xls")
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
inputfiles()
`bash report_run.bat`
getinformation()
`mkdir 01_datfiles`
`mkdir 02_csvfiles`
`mv loc_*.dat 01_datfiles`
csv2excel()
`mv *.csv 02_csvfiles`
`rm *.bat`
