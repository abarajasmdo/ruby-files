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
def inputfiles()
  h=File.new("mxlife.lock","w+")
  h.close
   if FileTest.exist?("TPS_cases.dat") == true
      a=File.read("TPS_cases.dat").squeeze(" ").split("\n")
      b=Array.new
      i=0
      a.each do |data|
        if data==" " then
        else
          b[i]=Array.new
          b[i]=data
          i=i+1
        end
      end
  end
  c=`ls *.37`.split[0].to_s
  g=File.new("tps_life_run.bat","w+")
  i=0
  b.each do |line|
  g<<"siesta_lite mxlife<< log
"<<c<<"

tabl
1
echo
"<<line<<"
all



fini
log
perl mxlife_output.pl
perl mxlife_stress.pl
rm mxlife_output.pl
rm mxlife_stress.pl
mv f50.dat "<< i <<"_mxlife_f50.dat
mv f45.dat "<< i <<"_mxlife_f45.dat
mv f48.dat "<< i <<"_mxlife_f48.dat
mv f59.dat "<< i <<"_mxlife_f59.dat
mv mxlife.csv "<< i <<"_mxlife_csv.dat
mv mxlife.f32 "<< i <<"_mxlife.f32
mv mxlife_output.xls "<< i <<"_mxlife_output.xls
mv mxlife_stress.xls "<< i <<"_mxlife_stress.xls
"
  i=i+1
  end
  g<<"rm mxlife.lock
"
  g.close
end
def getinformation()
  files=`ls *f50.dat`.split
  hfx=0
  files.each do |f|
    ff=f.gsub("_mxlife_f50.dat","").to_i
    if ff<10 then
      fff="00"+ff.to_s
    elsif ff<100 then
      fff="0"+ff.to_s
    elsif ff<1000 then
      fff=+ff.to_s
    end
    i=1
    File.read(f).gsub("TIME","TIME,").gsub("TEMPERATURE","TEMPERATURE,").gsub("STRESS","STRESS,").gsub("       ENTITY",",,ENTITY").gsub("                     ",",,,").split("STRESS, MULTIPLIER TABLE").each do|loc|
      if i==1 then
      else
        hlabl=loc.squeeze(" ").split[0].to_s
        h=File.new(fff+"_loc_"+hlabl+".csv","w+")
        j=1
        loc.squeeze(" ").split("\n").each do|lines|
          if lines[0..1] == " $" then
          elsif j==1 then
          elsif lines=="" then
          elsif lines[0..3] == " THE" then
            j=0
          else
            h<<lines.gsub(" ",",")<<"\n"
          end
          j=j+1
        end
        h.close        
      end
      i=i+1
    end
    csv2excel(fff,hfx)
    `rm *.csv`
    hfx=hfx+1
  end
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
  b.reverse_each do |sheet|
    if i>0 then
      sht = workbook.Worksheets.Add()
      sht.Name = "loc_"+sheet.split("_")[2].to_s.gsub(".csv","")
    else
      workbook.Sheets('Sheet1').Select
      sht = workbook.ActiveSheet
      sht.Name = "loc_"+sheet.split("_")[2].to_s.gsub(".csv","")
    end      
    wtemporal=excel.workbooks.open(Fld+"/"+sheet)
    wtemporal.ActiveSheet.Range("A1:Z100").Copy
    sht.Paste
    wtemporal.Close(0)
    i=1
  end
  workbook.SaveAs(Fld+"/"+labl+"_life.xls")
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
io=IO.popen("cygwin","w")
sleep(2)
io.write("cd "+Fld+"\n")
sleep(2)
io.write("bash tps_life_run.bat\n")
sleep(2)
io.write("exit"+"\n")
io.close
while FileTest.exist?(Fld+"/mxlife.lock") == true
  sleep 4
end
getinformation()
`rm *.bat`
`mkdir 03_lifefiles`
`mv *_mxlife*.* 03_lifefiles`
`mkdir 04_lifesum`
`mv *life.xls 04_lifesum`
