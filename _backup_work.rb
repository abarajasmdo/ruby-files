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

  This is Version I.0 to update Work Folder with the PPT information.
  If you encounter any error, please call me.
       Antonio Barajas (8*879-0611).
       
  Program build on Nov 2010
"
end
def getroot()
  folder=`cd`.split("\n")[0].gsub("\\","/")       #OBTENER LA RAIZ DE LA CARPETA DE TRABAJO
  return folder
end
def getfolders(root)
  a=`ls -l #{root}`.split("\n")
  c=Array.new
  a.each do |b|
      if b.split[0][0,1] == "d" then
        c << b.split[8]
      end
    end 
  return c
end
require 'win32ole'
require 'watir'
intro()
Fld=getroot()
A="D:/Work"
B="R:/servicios-activos/AEDS-00953_F101-F110-F118_LPT_Manufacturing_Support/001_MRB_Support/3_Work"
Folders=getfolders(A)
Folders.each do |foldertocopy|
  if foldertocopy[0,2] == "F1"  then
    f=`ls #{A}/#{foldertocopy}/*.ppt`.split[0]
    g=f.gsub(A,B)
    h=f.split("/")[3]
    i=g.gsub("/"+h,"")
    `cp #{f} #{g}`
  end
end
