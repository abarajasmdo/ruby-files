  #~ require 'hpricot'
require 'win32ole'
require 'watir'

a=String.new
a=`ls *.rst*`
b=Array.new
i=0
a.each_line do |lines|
  b[i]=lines.chomp!
  c=File.open("run.bat","w")
  c<<"siesta_lite <<idy\n"<<"gapf\n"<<b[i]<<"\n\n"<<"t\n"<<"-1\n"<<"idy"
  c.close
  `bash run.bat`
  d=b[i].gsub(".rst",".csv")
  puts "Gap File Created:"
  puts d
  puts ""
  i=i+1
end

if FileTest.exist?("run.bat")==true then
  `rm run.bat`
end

if FileTest.exist?("file.bat")==true then
  `rm file.bat`
end

if FileTest.exist?("file.lock")==true then
  `rm file.lock`
end
