begin

application = WIN32OLE.new('Excel.Application')



workbook_list=[]

for file_path in ["C:\\Users\\" + ENV['USERNAME'] + "\\Documents\\test.xlsx",Sketchup.find_support_file('Plugins')+'\\a\\test.xlsx',"C:\\Program Files (x86)\\Google\\Google SketchUp 8\\Plugins\\a\\test.xlsx"]
  if File.exist?(file_path)
    workbook_list.push([file_path,File.mtime(file_path)])
  end
end

workbook_list.sort!{|x,y|y[1]<=>x[1]}


workbook = application.Workbooks.Open(workbook_list[0][0])

worksheet=workbook.Worksheets("Sheet2")
            
worksheet.Range("A200:Z250").Clear

i = 1
$window_data.each{|definition|
	definition.instances.each{|instance|
		i = i + 1
		#$window_data.push([definition,@above,@height,@width,@type,@color,@grid,"drag"])
		a = instance.transformation.origin
		worksheet.Cells(202,i).Value = a[2].to_i
		worksheet.Cells(203,i).Value = definition.get_attribute("window","height",36)
		worksheet.Cells(204,i).Value = definition.get_attribute("window","width",48)
		


		if definition.get_attribute("window","from") == "spreedsheet"
			worksheet.Cells(205,i).Value = definition.get_attribute("window","type")
            if definition.get_attribute("window","grid") == 0
                worksheet.Cells(207,i).Value = "No"
            else
                worksheet.Cells(207,i).Value = "Yes"
            end
        else

			if (type = definition.get_attribute("window","type",0))==0
				worksheet.Cells(205,i).Value = "Verticle"
			elsif type == 1
				worksheet.Cells(205,i).Value = "Slider"
			else
	            worksheet.Cells(205,i).Value = "Fixed"
	        end
	        worksheet.Cells(207,i).Value = definition.get_attribute("window","grid",0)	
    	end
 

		if a[0] == 0
            worksheet.Cells(200,i).Value = "EW1"
            worksheet.Cells(201,i).Value = $width - a[1].to_i
        elsif a[0] == $length
        	worksheet.Cells(200,i).Value = "EW2"
        	if definition.get_attribute("window","from") == "spreedsheet"
        		worksheet.Cells(201,i).Value = a[1].to_i - definition.get_attribute("window","width").to_i
        	else
				worksheet.Cells(201,i).Value = a[1].to_i
			end
        elsif  a[1] == 0
            worksheet.Cells(200,i).Value = "SW1"
            worksheet.Cells(201,i).Value = a[0].to_i
        else
            worksheet.Cells(200,i).Value = "SW2"        
            worksheet.Cells(201,i).Value = $length - a[0].to_i
        end


	}
}

worksheet.Cells(208,1).Value = i - 1

rescue NoMemoryError

ensure
	workbook.save
	application.Workbooks.Close
	application.quit
end