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

  This is Version I.0 to modify MCASES with UIF file.
  If you encounter a real error, please call me.
       Antonio Barajas (8*879-0611).
       
  Program build on December 2010
"
end
def getfolder()
  folder=`cd`.split("\n")[0].gsub("\\","/")       #OBTENER LA RAIZ DE LA CARPETA DE TRABAJO
  return folder
end
def inputfiles()
  flock=File.new("_modif_mcas.lock","w+")
  flock.close
  
  fmcase=`ls *.37`.split("\n")[0]
  fuif=`ls *.uif`.split("\n")[0]
  
  frun=File.new("_modif_mcas.bat","w+")  

  frun<<"siesta_lite mcas<< log
"+fmcase+"

u
tabl
1 -1000

q
1
5
"+fuif+"

q
log
rm f*.dat
rm _modif_mcas.lock
clear
"
  frun.close
end
require 'watir'
intro()
Fld=getfolder()
inputfiles()
io=IO.popen("cygwin","w")
sleep(2)
io.write("cd "+Fld+"\n")
sleep(2)
io.write("bash _modif_mcas.bat\n")
sleep(2)
while FileTest.exist?(Fld+"/_modif_mcas.lock") == true
  sleep 4
end
`rm *.bat`
io.write("exit"+"\n")
io.close

