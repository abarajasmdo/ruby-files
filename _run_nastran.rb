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

  This is Version I.0 for run NASTRAN analysis.
  If you encounter a real error, please call me.
       Antonio Barajas (8*879-0611).
       
  Program build on October 2010
"
end
def getfolder()
  folder=`cd`.split("\n")[0].gsub("\\","/")       #OBTENER LA RAIZ DE LA CARPETA DE TRABAJO
  return folder
end
require 'win32ole'
require 'watir'
intro()
Fld=getfolder()
file=`ls *.dat`.split("\n")[0].to_s
g=File.new("run_nastran.bat","w+")
g<<"nastran2007r1 "<<file<<" mem=200MW sdir='"<<Fld<<"'\n"
g.close
io=IO.popen("cygwin","w")
sleep(2)
io.write("cd "+Fld+"\n")
sleep(2)
io.write("bash run_nastran.bat\n")
sleep(2)
io.write("exit"+"\n")
io.close
`rm run_nastran.bat`
