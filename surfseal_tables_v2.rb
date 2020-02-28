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

  This is Version 2.0 of Surfseal Automated Process.
  If you encounter a real error, please call me.
       Antonio Barajas (8*879-0611).
       
  Program release on April 2009
"
end
def getfolder()
  folder=`cd`.split("\n")[0].gsub("\\","/")       #OBTENER LA RAIZ DE LA CARPETA DE TRABAJO
  #~ if folder.split("/")[1].size > 1 then
      #~ drive=`vol`.split[3] + ":"
      #~ folder=drive+folder
  #~ else
    #~ if folder=~/^\/(\w+)\// then
      #~ drive=$1+":/"
    #~ end
    #~ folder[0,3]=drive
  #~ end
  return folder
end
def createinputfile()
  fend=File.new("fexit.dat","w+")
  fend<<"mkdir 04_data\n"
  fend<<"mv *.xls 04_data\n"
  fend<<"cp //quedatap02/aeds/aeae/SIXSIGMA/Six_Sigma_AMTO/Antonio_Barajas/2009_Engagement/surfseal_postprocess.xls " + Fld + "\n"
  fend.close
  h=File.new("surf.lock","w+")
  h.close
  f=File.new("surfseal_inp_list.dat","w+")
  f<<`ls *.rst`.gsub('.rst', '')
  f.close
  g=File.new("surfseal_run.bat","w+")
  g<<"mkdir 01_resume
mkdir 02_stress
mkdir 03_mcase
while read jobname1
do
echo $jobname1
siesta_lite << log
xsurf
mcas,on
dbgen,on
rrst,${jobname1}.rst,
"
  g<<File.read("surfseal_inp_patches.dat")
  g<<"fini
log
mv file18.dat surfseal_${jobname1}.out
mv file21.dat surfseal_${jobname1}.roll
mv file19.dat surfseal_${jobname1}.seal_geom
mv file29.dat surfseal_${jobname1}.max
mv file26.dat surfseal_${jobname1}.stress
mv file31.dat surfseal_${jobname1}.uif
mv file32.dat surfseal_${jobname1}.uof
mv *.out 01_resume
mv *.roll 01_resume
mv *.seal_geom 01_resume
mv *.max 01_resume
mv *.uif *.uof 03_mcase
mv *.stress 02_stress
rm file27.dat file30.dat
mv file*.dat 01_resume

done < surfseal_inp_list.dat
rm surf.lock
"
  g.close
end

def runsurfseal()
 `bash surfseal_run.bat`
  while FileTest.exist?("surf.lock") == true
    sleep 4
  end
  `rm *.bat`
end
def getstressdata()
  b=Hash.new
  stressfiles = `ls -G 02_stress`.split(".stress\n")
  stressfiles.each do |stress|
    a=File.read(Fld+"/02_stress/"+stress+".stress").split("####################################################################################################################################\n")
    b[stress]=Hash.new
    n=0
    a.each do |patch|
      if n>0 then
        pnam=patch.split("\n")[0].to_s
        b[stress][pnam]=patch.gsub("LEFT EDGE","LEFTEDGE").gsub("RIGH EDGE","RIGHEDGE").gsub("TOP  EDGE","TOPEDGE").gsub("BOTT EDGE","BOTTEDGE").gsub("SURF CENTRD","SURFCENTRD").gsub("LEFT TOP  CORNER","LEFTTOPCORNER").gsub("RIGH TOP  CORNER","RIGHTOPCORNER").gsub("LEFT BOTT CORNER","LEFTBOTTCORNER").gsub("RIGH BOTT CORNER","RIGHBOTTCORNER").gsub("     XDIR: -","").squeeze(" ").gsub(" ",",").gsub("\n,","\n").split("\n")
        b[stress][pnam][1]=patch.squeeze(" ").split("\n")[1]
        b[stress][pnam][8]="3D ELEM,SIGX,SIGY,TXY,SIG1,SIG2,SIGE,TEMP,COL,ROW,LOCATION"
        i=7
        while i>=0
          if i==1 then
          else
            b[stress][pnam].delete_at(i)
          end
          i=i-1
        end
      end
      n=1
    end
  end
  return b
end
def readidentifier()
  h=File.open("surfseal_inp_identifier.dat")
  hh=Hash.new
  h.read.split("\n").each do|line|
    i=line.split("\t")[1]
    if i=="TIME" then
    else
      if i.to_f<0 then
        i="      "+i
      else
        if i.to_f<10 then
          i="     "+i
        else
          if i.to_f<100 then
            i="    "+i
          else
            if i.to_f<1000 then
              i="   "+i
            else
              if i.to_f<10000 then
                i="  "+i
              else
                if i.to_f<100000 then
                  i=" "+i
                end
              end
            end
          end
        end
      end
      hh[i]=Array.new
      hh[i]=line.split("\t")
    end
  end
  hhh=Array.new
  i=0
  hh.keys.sort.each do|order|
    hhh[i]=Array.new
    hhh[i]=hh[order]
    i=i+1
  end
  return hhh
end
def createmcasfile(h)
  gg=`ls -G 03_mcase`.split("\n")[0]
  ghgh=File.new(Fld+"/03_mcase/mcase.lock","w+")
  ghgh.close
  ggg=File.new(Fld+"/03_mcase/surfseal_mcase_input.uif","w+")
  ggg<<File.open(Fld+"/03_mcase/"+gg).read
  ggg<<"\n"
  ggg<<File.open("surfseal_inp_material.dat").read
  ggg.close
  g=File.new(Fld+"/03_mcase/surfseal_mcase_generator.txt","w+")
  g<<"MPRE
surfseal_mcase
UIFN
surfseal_mcase_input.uif
TITL 1
Mcase for Xsurf Results
FTYP
UOF
FLST
LCAS	LRDA	TIME	FILN	CIDT
"
numcase=0
h.each do |data|
  g<<"1\t1\t"+data[1]+"\tsurfseal_"+data[0]+".uof\t"+data[2]+"\n"
  numcase=numcase+1
end
g<<"
NSTR
8	1100	1200	1300
ESTR
2200	2300	2700	4300
MAXC
"
g<<(numcase+5).to_s
"
$
"
  g.close
  fg=File.new(Fld+"/mcase_run.bat","w+")
    fg<<"cd 03_mcase
siesta_lite mcas<<idy
.set
surfseal_mcase_generator.txt
idy
rm *.dat
rm mcase.lock
"
  fg.close
end
def runmcase()
 `bash mcase_run.bat`
  while FileTest.exist?(Fld+"/03_mcase/mcase.lock") == true
    sleep 4
  end
  `rm *.bat`
end
def runmxlife()
#-----------------SIGE RANGE
  lock=File.new(Fld+"/03_mcase/mxlife.lock","w+")
  lock.close
  mxlf=File.new(Fld+"/mxlife_run.bat","w+")
  mxlf<<"cd 03_mcase
siesta_lite mxlife<<idy
surfseal_mcase.37

patch
11
a

idy
mv f50.dat surfseal_range.dat
rm f*.dat
rm *.lock
"
  mxlf.close
  `bash mxlife_run.bat`
  while FileTest.exist?(Fld+"/03_mcase/mxlife.lock") == true
    sleep 4
  end
  `rm *.bat`  
#-----------------MIN LIFE
  lock=File.new(Fld+"/03_mcase/mxlife.lock","w+")
  lock.close
  mxlf=File.new(Fld+"/mxlife_run.bat","w+")
  mxlf<<"cd "+Fld+"/03_mcase
siesta_lite mxlife<<idy
surfseal_mcase.37

patch
10
a

idy
mv f50.dat surfseal_life.dat
rm f*.dat
rm *.csv
rm *.xls
rm *.lock
"
  mxlf.close
#ABRIR CYGWIN
  IO.popen("cygwin", "r+") do |io|
    io.write("cd " + Fld+ "\n")
    io.write("bash mxlife_run.bat"+ "\n")
    io.write("exit"+ "\n")
    io.close_write
  end
  while FileTest.exist?(Fld+"/03_mcase/mxlife.lock") == true
    sleep 4
  end
  `rm *.bat`
end
def getsurfdata()
  filecrit=File.open("surfseal_crit_loc.xls","w+")
  filecrit<<"\tMIN LIFE\tMIN LIFE\tSIGE RANGE\tSIGE RANGE\tMAX SIG1\tMAXSIG1\n"
  filecrit<<"PATCH NAME\t2D ELEM\tLOC\t2D ELEM\tLOC\t2D ELEM\tLOC\n"
  c=Hash.new
  surffiles = `ls -G 03_mcase`.split("\n")
  surffiles.each do |surf|
    if surf =~/\.dat$/ then
      a=File.read(Fld+"/03_mcase/"+surf).split("\n       ENTITY  LOC")
      surfname=surf.gsub(".dat","")
      c[surfname]=Hash.new  
      n=0
      a.each do |patch|
        if n>0 then
          pnam=patch.split("\n     SURFSEAL ")[1].split("\n")[0].to_s.strip.sub(/(PATCH     )/, 'Region')
          c[surfname][pnam]=Array.new
          c[surfname][pnam]=patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").split("\n,")[0].to_s.split("EL2D,")
          c[surfname][pnam].delete("")
          c[surfname][pnam][0]="EL2D,EL2D LOC,SIGE RANGE,MAX TIME,MAX TEMP,MAX S11,MAX S22,MAX S33,MAX S12,MAX S23,MAX S13,EL3D,EL3D LOC,MIN TIME,MIN TEMP,MIN S11,MIN S22,MIN S33,MIN S12,MIN S23,MIN S13"
          if surfname=="surfseal_life" then
            c[surfname][pnam][0]="EL2D,EL2D LOC,CALC LIFE,MIN TIME,MAX TIME,MIN TEMP,MAX TEMP,MIN STRESS,MAX STRESS,RATIO,WALKER EXP,MISSION MIX,WALKER SALT,,,EL3D,EL3D LOC,M CALC LIFE,M MIN TIME,M MAX TIME,M MIN TEMP,M MAX TEMP,M MIN STRESS,M MAX STRESS,M RATIO,M WALKER EXP,M DAMAGE"
            critmlf = patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").split("\n,")[1].split("EL2D,")[1].split(",")[0] + "\t" + patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").split("\n,")[1].split("EL2D,")[1].split(",")[2]
            critser =patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").split("\n,")[2].split("EL2D,")[1].split(",")[0] + "\t" + patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").split("\n,")[2].split("EL2D,")[1].split(",")[2]
            critms1 = patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").split("\n,")[3].split("EL2D,")[1].split(",")[0] + "\t" + patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").split("\n,")[3].split("EL2D,")[1].split(",")[2]
            filecrit<<pnam<<"\t"<<critmlf<<"\t"<<critser<<"\t"<<critms1<<"\n"
          end
        end
        n=1
      end
    end
  end
  filecrit.close          
  return c
end
def array2excel(b)
  excel = WIN32OLE.new("excel.application")
  excel.DisplayAlerts = false
  excel.Interactive = false
  excel.ScreenUpdating = false
  if Iexcel==0 then
    WIN32OLE.const_load(excel, EC)
  end
  excel['Visible'] = false
  b.keys.each do |book|
    workbook = excel.Workbooks.Add(EC::XlWBATWorksheet)
    i=0
    b[book].keys.sort.reverse.each do |sht|
      if i>0 then
        sheet = workbook.Worksheets.Add()
        sheet.Name = sht
      else
        workbook.Sheets('Sheet1').Select
        sheet = workbook.ActiveSheet
        sheet.Name = sht
      end
      jj=1
      filedom=File.new("temporal.csv","w+")
      b[book][sht].each do |line|
        filedom<<line+"\n"
      end
      filedom.close
      wtemporal=excel.workbooks.open(Fld+"/temporal.csv")
      wtemporal.ActiveSheet.Range("A1:AZ5000").Copy
      sheet.Paste
      wtemporal.Close(0)
      `rm temporal.csv`
      i=1
    end
    workbook.SaveAs(Fld+"/"+book+".xls")
    excel.Interactive = true
    excel.ScreenUpdating = true
    excel.ActiveWorkbook.Close(0)
  end
  excel.Quit();
end
require 'hpricot'
require 'win32ole'
require 'watir'
Iexcel=0
intro()
Fld=getfolder()
createinputfile()
runsurfseal()
b=getstressdata()
array2excel(b)
h=readidentifier()
createmcasfile(h)
runmcase()
runmxlife()
c=getsurfdata()
Iexcel=Iexcel+1
array2excel(c)
`bash fexit.dat`
`rm fexit.dat`
#~ p Fld