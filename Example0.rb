  require 'watir'
  require 'win32ole'
  print "Hola Illya cuales son las medidas de la viga (Largo,Ancho) "
  vals = gets
  vals=vals.split(",")
  largo = vals[0].to_f
  ancho = vals[1].to_f
  
  
  archivo_1=File.open("Simple_Beam.inp","r+")
  lineas_archivo_1=archivo_1.read.split("\n")
  l_4= lineas_archivo_1[4].split(",")
  l_4[3] = largo
  l_4[4] = ancho
  lineas_archivo_1[4]=l_4.join(",")
  
  l_10= lineas_archivo_1[10].split(",")
  l_10[1]="node(" + largo.to_s
  l_10[2]=ancho
  lineas_archivo_1[10]=l_10.join(",")

  archivo_1.rewind
  archivo_1 << lineas_archivo_1.join("\n")
  archivo_1.close
  
  print "\nEmpieza el Analisis de Ansys"
a=`bash Ansys_Bash.sh`
  
  print "\nEmpieza post en Excel"
  archivo_1=File.open("Resul_Nodos.dat","r").read.split("\n")
  excel = WIN32OLE.new("excel.application")
  workbook = excel.Workbooks.Add(); 

  excel.visible=true
  workbook.sheets(1).cells(1,1).value="Nodo"
  workbook.sheets(1).cells(1,2).value="X"
  workbook.sheets(1).cells(1,3).value="Y"
  workbook.sheets(1).cells(1,4).value="UX"
  workbook.sheets(1).cells(1,5).value="UY"
  workbook.sheets(1).range("a1:e1").interior.ColorIndex = 40
  workbook.sheets(1).range("a1:e1").Font.Bold = 1
  workbook.sheets(1).range("a1:e1").HorizontalAlignment = 3

  j=2
  kolor=2
  archivo_1.each do |linea|
          linea=linea.split(",")
          	# Creates OLE object to Excel
          1.step(5,1) {|x|
                            workbook.sheets(1).cells(j,x).value =  linea[x-1]
                            if j%2 == 1 then
                              workbook.sheets(1).cells(j,x).interior.ColorIndex = 35
                              
                            else
                              workbook.sheets(1).cells(j,x).interior.ColorIndex = 2
                              
                            end 
                            
                          }
                          j +=1
	end
 
 #~ excelchart = workbook.Charts.Add()
 #~ workbook.Charts(1).charttype =72
 #~ workbook.Charts(1).SeriesCollection(1).xValues = "=Sheet1!R2C2:R65C2"
#~ workbook.Charts(1).SeriesCollection(1).Values = "=Sheet1!R2C4:R65C4"
#~ workbook.Charts(1).SeriesCollection.NewSeries
#~ workbook.Charts(1).SeriesCollection(2).xValues = "=Sheet1!R2C2:R65C2"
#~ workbook.Charts(1).SeriesCollection(2).Values = "=Sheet1!R2C5:R65C5"

print "\nCarga en Orion"
  ie = Watir::IE.start("http://orion.mtc.ge.com")
  if ie.contains_text("Orion") then
    ie.text_field(:name, "typedUsername").set("210039289")

    ie.text_field(:name, "typedPassword").set("***********")
    ie.button(:value, "Login").click
  end

 ie.goto("http://orion.mtc.ge.com/orion/modules/voucher/add.php")
 
 ie.text_field(:name, "usedHours").set("9")
 ie.text_field(:name, "description").set("Analisis Exhaustivo de una extremadamente compleja viga en cantiliver, analizada mediante metodos de elemento finito en ANSYS")
 