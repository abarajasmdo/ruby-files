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

       Antonio Barajas (8*879-0611).
"
end
def getfolder()
  folder=`cd`.split("\n")[0]				#OBTENER LA RAIZ DE LA CARPETA DE TRABAJO
  return folder
end
def postdata()
	a=File.read("CT7-8X_CC_components.txt").split("\n\n")
	b=File.new("CT7-8X_CC_Components.cdb","w+")
	i=0
  a.each do |comp|  
		if i== 0 then
		else
			comp.gsub("-"," -").split('*Pool "Brick"')[1].split.size
			compcount=comp.gsub("-"," -").split('*Pool "Brick"')[1].split.size
			head=comp.split('*Pool "Brick"')[0].gsub('*Group "','CMBLOCK,').gsub('" "element" "  0 255 255"',',ELEM,').gsub("\n","")<<compcount.to_s<<"\n(8i10)\n"
			j=0
			body=""
			comp.gsub("-"," -").split('*Pool "Brick"')[1].split.each do |member|
				if j==8 then
					j=0
					body=body<<"\n"<<member.rjust(10)
				else
					body=body<<member.rjust(10)
				end
				j=j+1
			end
			b<<head<<body<<"\n"
		end
		i=i+1
	end
	b.close
end
require 'watir'
intro()
Fld=getfolder()
postdata()
