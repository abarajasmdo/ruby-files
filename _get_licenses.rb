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

  This is Version I.0 for checking availability of the Licences used in ANSYS.
  If you encounter a real error, please call me.
       Antonio Barajas (8*879-0611).
       
  Program build on Sept 2010
"
end
def getfolder()
  folder=`cd`.split("\n")[0].gsub("\\","/")       #OBTENER LA RAIZ DE LA CARPETA DE TRABAJO
  return folder
end
def readdata()
  a=File.read("file.out").split("How licenses")[2].split("Checked out: ")
  i=0
  b=Hash.new
  a.each do |loadfile|
    if i==0 then
      i=1
    else
      if loadfile=="" then
      else
        username=loadfile.split(" ")[1].to_s
        licenses=loadfile.split(" - ")[1].to_s.split(" ")[0].to_i
        if username==nil then
        else
          if b[username]==nil then
          b[username]=Array.new
          b[username][0]=0
          end
          b[username][0]=b[username][0]+licenses
          b[username]<<licenses
        end
      end
    end
  end
  return b
end
#~ require 'hpricot'
require 'win32ole'
require 'watir'
intro()
Fld=getfolder()
h=readdata()
g=File.new("_used_licenses.dat","w+")
g<<"\tUSER\t\tLICENSE(S)\n"
grand_total=0
h.keys.sort.each do|data|
  if data.length>7 then
    texttab="\t"
  else
    texttab="\t\t"
  end
  data_total=h[data][0].to_s
  h[data].delete_at(0)
  if data_total.to_i>14 then
    mark="*"
  else
    mark=""
  end
  grand_total=grand_total+data_total.to_i
  g<<mark+"\t"+data+texttab+"TOTAL: "+data_total+"\t("+h[data].join(",")+")\n"
end
g<<"\n\n LICENCE TOTAL: "+grand_total.to_s
g.close
