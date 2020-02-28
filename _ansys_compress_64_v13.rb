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

  folder=`cd`.split("\n")[0].gsub("\\","/")            #OBTENER LA RAIZ DE LA CARPETA DE TRABAJO

  flist=`ls *.rst`.split("\n")
  
  f=File.open("_ansys_compress.dat","w+")
  flist.each do |fname|
    f<<"fini
/aux3
fileaux3,"+fname.gsub(".",",")+"
addf,"+fname.gsub(".",",")+"
delete,set,1
compress
fini
/clear,nostart

"
  end
  f.close

  file=File.open("_Ansys_Compress_64.sh","w+")
  file<<'export ANSYS130_DIR="V:\aeapps\Ansys\v13_0_x64\Server\v130\ansys"
export ANSYS_SYSDIR="winx64"
export ANSYS110_PRODUCT=ANE3FL
export ANSYSLI_SERVERS="28024@tnsp07282.ae.ge.com;28024@tnsp07281.ae.ge.com;28024@tnsp07280.ae.ge.com"
export ANSYSLMD_LICENSE_FILE="28020@tnsp07280.ae.ge.com;28020@tnsp07281.ae.ge.com;28020@tnsp07282.ae.ge.com"

"V:/aeapps/Ansys/v13_0_x64/Server/v130/ansys/bin/winx64/ANSYS130.exe"'+" -b -i _ansys_compress.dat -o file.out"
  file.close


`bash _Ansys_Compress_64.sh`

`rm _Ansys_Compress_64.sh`
`rm _ansys_compress.dat`
