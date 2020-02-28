require 'watir'
require 'hpricot'
require 'win32ole'
require 'tk'

def intro()
  puts " "
  puts " "
  puts " "
  puts "        ****************************************************************        "
  puts " "
  puts "        PLEASE DO NOT CLOSE THIS WINDOW OR TRY TO OPEN AN EXCEL DOCUMENT        "
  puts " "
  puts "        ****************************************************************        "
  puts " "
  puts "        ------------------ PLEASE WAIT SCRIPT RUNNING ------------------        "
  puts " "
  puts "        ****************************************************************        "
  puts "        ________________________________________________________________        "
  puts "                                       CREATED BY ANTONIO ZAMORA-NHCUMFT        "
  puts "                                                              8*743 0731        "
  puts "                                                                FEB 2009        "
end

def message()
  
  root_1 = TkRoot.new{ title "REPORT"}
  frame_5=TkFrame.new(root).pack(:fill=>:both)
  report = TkLabel.new(root_1, :text=>'BOTH LISTS ARE IDENTICAL', :font=>'Helvetica 11').pack
  report_1 = TkLabel.new(root_1, :text=>'NO CHANGES WERE FOUND', :font=>'Helvetica 11').pack
  ok = TkButton.new(root_1, :text=>'OK', :font=>'Helvetica 11', :command => proc{exit} ).pack
  root_1.geometry "+250+250"
  Tk.mainloop
  
end

def arguments()
  
  root=TkRoot.new{title "INPUT FILE"}
  win=TkToplevel.new{title "INPUT FILE"}
  win_1=TkToplevel.new{title "COLUMN DATE"}
  win_2=TkToplevel.new{title "COLUMN DATE"}
  win_3=TkToplevel.new{title "SAVE FILE AS"}
  
  win.withdraw
  win_1.withdraw
  win_2.withdraw
  win_3.withdraw
  
  frame = TkFrame.new(root).pack(:fill=>:both)
  frame_1=TkFrame.new(win).pack(:fill=>:both)
  frame_2=TkFrame.new(win_1).pack(:fill=>:both)
  frame_3=TkFrame.new(win_2).pack(:fill=>:both)
  frame_4=TkFrame.new(win_3).pack(:fill=>:both)

  label=TkLabel.new(frame, :text=>'ENTER THE FIRST FILE:', :font=>'Helvetica 11').pack
  text_1=TkEntry.new(frame).pack(:fill=>:x)
  button = TkButton.new(root, :text=>'OK', :font=>'Helvetica 11', :command=>proc{@arg_1=text_1.value; root.withdraw; win.deiconify}).pack
  button_1=TkButton.new(root, :text=>'CANCEL', :font=>'Helvetica 11', :command=>proc{exit}).pack

  label=TkLabel.new(frame_1, :text=>'ENTER THE SECOND FILE:', :font=>'Helvetica 11').pack
  text_2=TkEntry.new(frame_1).pack(:fill=>:x)
  button = TkButton.new(win, :text=>'OK', :font=>'Helvetica 11', :command=>proc{@arg_2=text_2.value; win.withdraw; win_1.deiconify}).pack
  button_1 = TkButton.new(win, :text=>'CANCEL', :font=>'Helvetica 11', :command=>proc{exit}).pack
  
  label=TkLabel.new(frame_2, :text=>'ENTER THE DATE OF THE FIRST FILE:', :font=>'Helvetica 11').pack
  text_3=TkEntry.new(frame_2).pack(:fill=>:x)
  button = TkButton.new(win_1, :text=>'OK', :font=>'Helvetica 11', :command=>proc{@arg_3=text_3.value; win_1.withdraw; win_2.deiconify}).pack
  button_1 = TkButton.new(win_1, :text=>'CANCEL', :font=>'Helvetica 11', :command=>proc{exit}).pack
  
  label=TkLabel.new(frame_3, :text=>'ENTER THE DATE OF THE SECOND FILE:', :font=>'Helvetica 11').pack
  text_4=TkEntry.new(frame_3).pack(:fill=>:x)
  button = TkButton.new(win_2, :text=>'OK', :font=>'Helvetica 11', :command=>proc{@arg_4=text_4.value; win_2.withdraw; win_3.deiconify}).pack
  button_1 = TkButton.new(win_2, :text=>'CANCEL', :font=>'Helvetica 11', :command=>proc{exit}).pack
  
  label=TkLabel.new(frame_4, :text=>'SAVE FILE AS:', :font=>'Helvetica 11').pack
  text_5=TkEntry.new(frame_4).pack(:fill=>:x)
  button = TkButton.new(win_3, :text=>'OK', :font=>'Helvetica 11', :command=>proc{@arg_5=text_5.value; Tk.root.destroy}).pack
  button_1 = TkButton.new(win_3, :text=>'CANCEL', :font=>'Helvetica 11', :command=>proc{exit}).pack
  
  root.geometry "+250+250"
  win.geometry "+250+250"
  win_1.geometry "+250+250"
  win_2.geometry "+250+250"
  win_3.geometry "+250+250"
  
  Tk.mainloop
  Tk.restart
  
  data(@arg_1, @arg_2, @arg_3, @arg_4, @arg_5)
  
end

module EC
end

def data(arg_1, arg_2, arg_3, arg_4, arg_5)
  
  excel=WIN32OLE.new("excel.application")
  WIN32OLE.const_load(excel, EC) 
  excel.DisplayAlerts=false
  excel.Visible=false
  workbook=excel.Workbooks.open(Fld+"/"+arg_1) #NAME OF THE FILE
  #~ sheet=workbook.Sheets('ALL PARTS').Select; #SELECTION OF THE SHEET TO STRACT THE INFORMATION IF THE FILE HAS MORE THAN ONE SHEET
  sheet=workbook.ActiveSheet
  
  line = '2'
  data = []
  col=[]
  while sheet.Range("a#{line}")['Value']
    data << sheet.Range("a#{line}:d#{line}")['Value'].flatten!
    col<<sheet.Range("a#{line}:b#{line}")['Value']
    line.succ!
  end
  
  excel.ActiveWorkbook.Close(0)
  
  workbook=excel.Workbooks.open(Fld+"/"+arg_2) #NAME OF THE FILE TO BE COMPARED
  #~ sheet=workbook.Sheets('17100-Q0A').Select; #SELECTION OF THE SHEET TO STRACT THE INFORMATION IF THE FILE HAS MORE THAN ONE SHEET
  sheet=workbook.ActiveSheet
  line='2'
  data_1=[]
  col_1=[]
  while sheet.Range("a#{line}")['Value']
    data_1<<sheet.Range("a#{line}:d#{line}")['Value'].flatten!
    col_1<<sheet.Range("a#{line}:b#{line}")['Value']
    line.succ!
  end
  
  excel.ActiveWorkbook.Close(0)
  excel.Interactive = true
  excel.ScreenUpdating = true
  excel.Quit
  
  if col.length >=col_1.length
    
    title=[]
    title=arg_3 #TITLE OF THE FIRST FILE TO BE INTRODUCED IN THE TABLE
    
    title_1=[]
    title_1=arg_4 #TITLE OF THE SECOND FILE TO BE INTRODUCED IN THE TABLE
    
    i=0
    c=[]
    col.each do |dat|
      if col_1.include?(dat)
      else
        c<<data[i]
        c.compact!
      end
      i=i+1
    end
  
    j=0
    l=0
    d=[]
    col.each do |dat|
      if col.include?(col_1[l])
      else
        d<<data_1[j]
        d.compact!
      end
      l=l+1
      j=j+1
    end

    k=0
    e=[]
    data.each do |part|
      data_1.each do |list|
        if part.include?(list[0])
          if part.include?(list[1])
            if part.include?(list[3])
            else
              e<<data[k]
              e.compact!
            end
          end
        end
      end
      k=k+1
    end
    
    h=0
    f=[]
    data_1.each do |parte|
      data.each do |lista|
        if lista.include?(parte[0])
          if lista.include?(parte[1])
            if lista.include?(parte[3])
            else
              f<<data_1[h]
              f.compact!
            end
          end
        end
      end
      h=h+1
    end
  else
    
    title=[]
    title=arg_4 #TITLE OF THE SECOND FILE TO BE INTRODUCED IN THE TABLE
    
    title_1=[]
    title_1=arg_3 #TITLE OF THE FIRST FILE TO BE INTRODUCED IN THE TABLE
    
    i=0
    c=[]
    col_1.each do |dat|
      if col.include?(dat)
      else
        c<<data_1[i]
        c.compact!
      end
      i=i+1
    end
  
    j=0
    l=0
    d=[]
    col_1.each do |dat|
      if col_1.include?(col_1[l])
      else
        d<<data[j]
        d.compact!
      end
      l=l+1
      j=j+1
    end

    k=0
    e=[]
    data_1.each do |part|
      data.each do |list|
        if part.include?(list[0])
          if part.include?(list[1])
            if part.include?(list[3])
            else
              e<<data_1[k]
              e.compact!
            end
          end
        end
      end
      k=k+1
    end

    h=0
    f=[]
    data.each do |parte|
      data_1.each do |lista|
        if lista.include?(parte[0])
          if lista.include?(parte[1])
            if lista.include?(parte[3])
            else
              f<<data[h]
              f.compact!
            end
          end
        end
      end
      h=h+1
    end
  end
  
  summary(title,title_1,c,d,e,f,arg_5)
  
end
  
def summary(title,title_1,c,d,e,f,arg_5)
  
  excel=WIN32OLE.new("excel.application")
  excel.DisplayAlerts=false
  excel.Visible=false
  excel.Interactive = false
  excel.ScreenUpdating = false
  workbook = excel.Workbooks.Add(EC::XlWBATWorksheet)
  sheet=workbook.ActiveSheet
  sheet.Name="Summary"
  
  sheet.Range("a1")['Value']="PPN"
  sheet.Range("a:a").EntireColumn.Autofit
  sheet.Range("b1")['Value']="Part"
  sheet.Range("b:b").EntireColumn.Autofit
  sheet.Range("c1")['Value']="Nomenclature"
  sheet.Range("c:c").EntireColumn.Autofit
  sheet.Range("d1")['Value']=title
  sheet.Range("d:d").EntireColumn.Autofit
  sheet.Range("e1")['Value']=title_1
  sheet.Range("e:e").EntireColumn.Autofit
  sheet.Range("a1:e1").HorizontalAlignment = EC::XlCenter
  sheet.Range("a1:e1").interior.ColorIndex=37
  
  
  if c.empty?
  else
    line='2'
    c.each do|data|
      colum='a'
      data.each do |list|
        sheet.Range("#{colum}#{line}")['Value']="\'"+list
        sheet.Range("#{colum}#{line}").HorizontalAlignment = EC::XlCenter
        sheet.Range("#{colum}:#{colum}").EntireColumn.Autofit
        sheet.Range("e#{line}")['Value']="Not Included"
        sheet.Range("e#{line}").HorizontalAlignment = EC::XlCenter
        sheet.Range("e:e").EntireColumn.Autofit      
        colum.succ!
      end
      line.succ!
    end
  end
  
  if d.empty?
  else
    d.each do |data|
      colum='a'
      data.each do |list|
        sheet.Range("#{colum}#{line}")['Value']="\'"+list
        sheet.Range("#{colum}#{line}").HorizontalAlignment = EC::XlCenter
        sheet.Range("#{colum}:#{colum}").EntireColumn.Autofit
        sheet.Range("d#{line}")['Value']="Not Included"
        sheet.Range("d#{line}").HorizontalAlignment = EC::XlCenter
        sheet.Range("d:d").EntireColumn.Autofit
        colum.succ!
      end
      data[3].each do |list|
        sheet.Range("e#{line}")['Value']="\'"+list
        sheet.Range("e#{line}").HorizontalAlignment = EC::XlCenter
        sheet.Range("e:#{colum}").EntireColumn.Autofit
      end
      line.succ!
    end
  end
    
  line_1='2'
  while sheet.Range("a#{line_1}")['Value']
    line_1.succ!
  end
    
  if e.empty?
  else
    e.each do |data|
      colum='a'
      data.each do |list|
        sheet.Range("#{colum}#{line}")['Value']="\'"+list
        sheet.Range("#{colum}#{line}").HorizontalAlignment = EC::XlCenter
        sheet.Range("#{colum}:#{colum}").EntireColumn.Autofit
        colum.succ!
      end
      line.succ!
    end
  end
    
  if f.empty?
  else
    f.each do |data|
      data[3].each do |list|
        sheet.Range("e#{line_1}")['Value']="\'"+list
        sheet.Range("e#{line_1}").HorizontalAlignment = EC::XlCenter
        sheet.Range("e:e").EntireColumn.Autofit
      end
      line_1.succ!
    end
  end
    
  line_2='2'
  i=0
  while sheet.Range("a#{line_2}").offset(i,0).value
    sheet.Range("a#{line_2}:e#{line_2}").offset(i,0).Interior.ColorIndex = 35
    line_2.succ!
    i=i+1
  end
    
  line_3='2'
  if sheet.Range("a#{line_3}")['Value']==nil
    excel.ActiveWorkbook.Close(0)
    excel.Interactive = true
    excel.ScreenUpdating = true
    excel.Quit  
    message()
  else
    list=`ls`
    if list.include?("Summary_tables")==true
      workbook.SaveAs(Fld+"/Summary_tables"+"/"+arg_5) #NAME OF THE FILE WITH THE SUMMARY TABLE
      excel.ActiveWorkbook.Close(0)
      excel.Interactive = true
      excel.ScreenUpdating = true
      excel.Quit     
    else
      `md Summary_tables` #NAME OF THE NEW DIRECTORY WHERE THE FILES ARE TO BE SAVED
      sleep 5
      workbook.SaveAs(Fld+"/Summary_tables"+"/"+arg_5) #NAME OF THE FILE WITH THE SUMMARY TABLE
      excel.ActiveWorkbook.Close(0)
      excel.Interactive = true
      excel.ScreenUpdating = true
      excel.Quit     
    end
  end
end

intro()
raiz=`pwd`.split("\n")[0]       #OBTAIN THE ROOT OF THE ADDRESS DIRECTORY
drive=`vol`.split[3]                #OBTAIN THE WORKING DISK
Fld=drive+":"+raiz                #GENERATE THE WORKING ADDRESS

arguments()