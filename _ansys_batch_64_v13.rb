#~ require 'watir'
#~ require 'win32ole'
  puts " "
  puts " "
  puts " "
  puts "        RRRRRRRR    PPPPPPPP   EEEEEEEE        GGGGGGGG  EEEEEEEE  IIIIIIII  QQQQQQQQ  "
  puts "        RR     RR   PP     PP  EE              GG        EE           II     QQ    QQ  "
  puts "        RR     RR   PP     PP  EE              GG        EE           II     QQ    QQ  "
  puts "        RRRRRRRR    PPPPPPPP   EEEEEE          GGGGGGGG  EEEEEE       II     QQ    QQ  "
  puts "        RR    RR    PP         EE              GG    GG  EE           II     QQ    QQ  "
  puts "        RR     RR   PP         EE              GG    GG  EE           II     QQQQQQQQ  "
  puts "        RR      RR  PP         EEEEEEEE        GGGGGGGG  EEEEEEEE  IIIIIIII         QQ "
  puts ""
  puts "        Rotating Parts Engineering at GEIQ                                             "        
  puts "        ****************************************************************************** "
  puts " "
  puts "               PLEASE DO NOT CLOSE THIS WINDOW                                         "
  puts " "
  puts "        ****************************************************************************** "
  puts " "
  puts "        ------------------------- PLEASE WAIT SCRIPT RUNNING ------------------------- "
  puts " "
  puts "        ****************************************************************************** "
  puts "        ______________________________________________________________________________ "
  puts "                                                     CREATED BY ANTONIO ZAMORA-NHCUMFT "
  puts "                                                   MODIFIED BY ANTONIO BARAJAS-NHCUKHT "
  puts "                                                                            APRIL 2009 "
  
def input()
  b=`ls *.dat`
  puts""
  puts b
  puts""
  puts "      ENTER THE NAME OF THE '.dat' FILE:"
  @arg_1=gets.chomp!
  puts""
  puts"      ENTER THE NUMBER OF PROCESSORS:"
  @arg=gets.chomp!
  argument(@arg, @arg_1)
end

def argument(arg, arg_1)
  file=File.open("_Ansys_Bash_64.sh","w")
  file<<'export ANSYS130_DIR="V:\aeapps\Ansys\v13_0_x64\Server\v130\ansys"
export ANSYS_SYSDIR="winx64"
export ANSYS110_PRODUCT=ANE3FL
export ANSYSLI_SERVERS="28024@tnsp07282.ae.ge.com;28024@tnsp07281.ae.ge.com;28024@tnsp07280.ae.ge.com"
export ANSYSLMD_LICENSE_FILE="28020@tnsp07280.ae.ge.com;28020@tnsp07281.ae.ge.com;28020@tnsp07282.ae.ge.com"

"V:/aeapps/Ansys/v13_0_x64/Server/v130/ansys/bin/winx64/ANSYS130.exe"'+" -b -np #{arg} -i #{arg_1} -o file.out"
  file.close
end

input()

#~ while FileTest.exist?("lock.lock") == true
  #~ puts ""
  #~ puts "*********************"
  #~ puts "WAITING FOR PREVIOUS ANALYSIS TO FINISH!!!!"
  #~ puts "*********************"
  #~ puts ""
  #~ sleep 60
#~ end

a=`bash _Ansys_Bash_64.sh`

`rm _Ansys_Bash_64.sh`
