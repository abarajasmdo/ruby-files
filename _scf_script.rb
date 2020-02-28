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

  This is Version I.0 of SCF runs.
  If you encounter a real error, please call me.
       Antonio Barajas (8*879-0611).
       
  Program build on October 2010
"
end
def getfolder()
  folder=`cd`.split("\n")[0].gsub("\\","/")       #OBTENER LA RAIZ DE LA CARPETA DE TRABAJO
  return folder
end
def inputfiles()
  flock=File.new("sigcor.lock","w+")
  flock.close
  fadd=File.read("sigcor_addinfo.inp").gsub("\n"," \n")
  mc2d=File.read("sigcor_mcases.inp").split("\n")[0]
  mc3d=File.read("sigcor_mcases.inp").split("\n")[1]
  plock=File.new(mc2d.gsub(".37",".inp"),"w+")
  plock<<"noplot"
  plock.close
  tmf=fadd.split("\n")[0]
  scomp3d=fadd.split("\n")[1]
  scftyp=fadd.split("\n")[2]
  scfloc=fadd.split("\n")[4]
  tpoff=fadd.split("\n")[5]
  scoff=fadd.split("\n")[6]
  frun=File.new("sigcor_run.bat","w+")  
  frun<<"siesta_lite << log
sigcor
"+mc2d+"

"+mc3d+"

"+tpoff+"
"+scoff+"
tmf
"+tmf+"
"+scfloc+"

"+scomp3d+"\n"
  el2d=File.read("sigcor_elems.inp").split("\n")
  el2d.each do |line|
    scomp2d=fadd.split("\n")[3]
    scomp2d.split("\t").each do |comp|
frun<<scftyp+"
"+line+"




"+comp+"
n\n"
if fadd.split("\n")[7] == nil then
else
  fadd.split("\n")[7].split(" ").each do |targetlabl|
frun<<"target
"+targetlabl+"
n\n"
  end
end
frun<<"redo\n"
    end
  end
  frun<<"q
q
log
mv f58.dat sigcor_run.out
rm f45.dat
rm f59.dat
rm f33.dat
rm f49.dat
rm sigcor.lock
rm "+mc2d.gsub(".37",".inp")+"
clear
"
  frun.close
end
def postdata()
  files=`ls sigcor_run.out`.split
  files.each do |f|  
    a=File.read(f).split("  LOCATION:")
    i=0
    j=0
    b=Array.new
    a.each do |sigcor|
      if i==0 then
        i=1
      else
        elem3d=sigcor.split("3D Stress is SIG:")[1].split[0]
        target=sigcor.split("TARGET:")[1].split("\n")[0].squeeze(" ")
        comb2d=sigcor.split("The 2D element is:")[1].split[0]
        rsquare=sigcor.split("R^2  AVE_ERR  STDEV")[1].split[0]
        aveerr=sigcor.split("R^2  AVE_ERR  STDEV")[1].split[1]
        stdev=sigcor.split("R^2  AVE_ERR  STDEV")[1].split[2]
        lablkt=""
        if sigcor.include? "Kt HOOP is" then
          hoopkt=sigcor.split("Kt HOOP is")[1].split[0]
          lablkt=lablkt+"H"
        else
          hoopkt="0"
        end
        if sigcor.include? "Kt RADIAL is" then
          radialkt=sigcor.split("Kt RADIAL is")[1].split[0]
          lablkt=lablkt+"R"
        else
          radialkt="0"
        end
        if sigcor.include? "Kt AXIAL is" then
          axialkt=sigcor.split("Kt AXIAL is")[1].split[0]
          lablkt=lablkt+"A"
        else
          axialkt="0"
        end
        if sigcor.include? "Kt SHEAR is" then
          shearkt=sigcor.split("Kt SHEAR is")[1].split[0]
          lablkt=lablkt+"S"
        else
          shearkt="0"
        end
        if sigcor.include? "Kt GAP_FN_LIST_" then
          lablkt=lablkt+"G"          
          gapkt=sigcor.split("Kt GAP_FN_LIST_")[1].split[2]
          gapnm=sigcor.split("Kt GAP_FN_LIST_")[1].split[0]
          if sigcor.split("Kt GAP_FN_LIST_")[1].include? "Kt GAP_FN_LIST_" then
          gapkt2=sigcor.split("Kt GAP_FN_LIST_")[1].split("Kt GAP_FN_LIST_")[1].split[2]
          gapnm2=sigcor.split("Kt GAP_FN_LIST_")[1].split("Kt GAP_FN_LIST_")[1].split[0]            
          else
          gapkt2="0"
          gapnm2="0"
          end
        else
          gapkt="0"
          gapnm="0"
          gapkt2="0"
          gapnm2="0"
        end
        b[j]=elem3d+","+target+","+comb2d+","+lablkt+","+hoopkt+","+radialkt+","+axialkt+","+shearkt+","+gapnm+","+gapkt+","+gapnm2+","+gapkt2+","+aveerr+","+stdev+","+rsquare
        j=j+1
      end
    end
    g=File.new("sigcor_summary_"+File.read("sigcor_addinfo.inp").split("\n")[2]+".csv","w+")
    g<<"3DELEM,TARGET,2DELEM,KT COMPONENTS,HOOP KT,RADIAL KT,AXIAL KT,SHEAR KT,GAP NAME,GAP KT,GAP2 NAME,GAP2 KT,AVE ERR,STD DEV,R2\n"
    b.each do|data|
      g<<data+"\n"
    end
    g.close
  end
  grun=File.new("changename.bat","w+")  
  grun<<"mv sigcor_run.out sigcor_summary_"+File.read("sigcor_addinfo.inp").split("\n")[2]+".out\n"
  grun.close
  `changename.bat`
end

require 'watir'
intro()
Fld=getfolder()
inputfiles()
io=IO.popen("cygwin","w")
sleep(2)
io.write("cd "+Fld+"\n")
sleep(2)
io.write("bash sigcor_run.bat\n")
sleep(2)
while FileTest.exist?(Fld+"/sigcor.lock") == true
  sleep 4
end
io.write("exit"+"\n")
io.close
postdata()
`rm *.bat`
