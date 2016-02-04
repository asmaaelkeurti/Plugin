entities = Sketchup.active_model.entities

for i in 0..entities.length
  if(entities[0] != nil)
      entities[0].erase!
  end
end

application = WIN32OLE.new('Excel.Application')
if !File.directory?("C:\\Program Files (x86)\\Google\\Google SketchUp 8\\Plugins\\a")
  workbook = application.Workbooks.Open("C:\\Users\\"+ENV['USERNAME']+"\\AppData\\Roaming\\SketchUp\\SketchUp 2016\\SketchUp\\Plugins\\a\\test")
else
  workbook = application.Workbooks.Open("C:\\Program Files (x86)\\Google\\Google SketchUp 8\\Plugins\\a\\test")
end
worksheet=workbook.Worksheets("Sheet2")


begin

$a = []

opening = []
for i in 1..12
  if worksheet.Cells(28,1+i).Value == "EW2"
    opening = opening + [[worksheet.Cells(24,1+i).Value.to_f/12,worksheet.Cells(26,1+i).Value.to_f/12,worksheet.Cells(25,1+i).Value.to_f/12]]
  end
  if worksheet.Cells(52, 1+i).Value == "EW2"
    opening = opening + [[worksheet.Cells(53, 1+i).Value.to_f/12,worksheet.Cells(48,1+i).Value.to_f/12,worksheet.Cells(49,1+i).Value.to_f/12]]
  end
end


#opening = [[24,10,10]] #offset width,height

heel = worksheet.Cells(502,2).Value
heel = heel[0..heel.index("'")].to_f-0.5


wcht = worksheet.Cells(10,2).Value.to_f/12.0
width = worksheet.Cells(2,2).Value.to_f/12.0
height = (worksheet.Cells(4,2).Value.to_f+4+heel+3.75-5.5+2-5.5-worksheet.Cells(505,2).Value.to_f)/12.0 
pitch = worksheet.Cells(5,2).Value.to_f
x = 3-(3-width%3)%3*0.5


while x <  width
  if x <= width/2.0
    $a = $a + [[[x*12,height*12+x*pitch,0],[x*12,0,0]]]
    x = x + 3
  else 
    $a = $a + [[[x*12,2*(height*12+width*pitch/2)-height*12-x*pitch,0],[x*12,0,0]]]
    x = x + 3
  end
end  


entities.add_dimension_linear([0,0,0],[0,wcht*12,0],[-20,0,0]).arrow_type = Sketchup::Dimension::ARROW_SLASH if wcht > 1
center = [width*12/2.0,wcht*12,0]
centerR = [width*12/2.0,wcht*12,0]

entities.add_line([0,0,0],[width*12,0,0])
entities.add_line([0,wcht*12,0],[width*12,wcht*12,0])
for j in opening
	entities.add_line([j[0]*12,wcht*12,0],[j[0]*12+j[1]*12,wcht*12,0]).erase!
  entities.add_line([j[0]*12,0,0],[j[0]*12+j[1]*12,0,0]).erase!
	for i in $a
		if i[1][0]>=j[0]*12 and i[1][0]<=(j[0]+j[1])*12
      if (i[1][0]<=width*12/2.0 and i[1][0]-j[0]*12<3*12) 
                i[1][1]=wcht*12
      elsif (i[1][0]>width*12/2.0 and j[0]*12+j[1]*12-i[1][0]<3*12) 
        i[1][1]=wcht*12
      else
        i[1][1]=j[2]*12
      end
		end
	end
	entities.add_line([j[0]*12,0,0],[j[0]*12,j[2]*12,0])
	entities.add_line([j[0]*12,j[2]*12,0],[(j[0]+j[1])*12,j[2]*12,0])
	entities.add_line([(j[0]+j[1])*12,j[2]*12,0],[(j[0]+j[1])*12,0,0])

  if center[0]-18>=j[0]*12 and center[0]+18<=j[0]*12 + j[1]*12
    center[1] = j[2]*12
  end

  if centerR[0]>=j[0]*12 and centerR[0]<=j[0]*12 + j[1]*12-3*12
    centerR[1] = j[2]*12
  end


end
if ((width-2*x)/3)%2==1
  entities.add_dimension_linear([width*12/2.0,height*12+width*pitch/2.0,0],center,[1,0,0]).arrow_type = Sketchup::Dimension::ARROW_SLASH
end
if ((width-2*x)/3)%2!=1
  entities.add_dimension_linear([width*12/2.0,height*12+width*pitch/2.0,0],centerR,[20,0,0]).arrow_type = Sketchup::Dimension::ARROW_SLASH
end


p1 = [0,0,0]
p2 = [0,height*12,0]
p3 = [width*6,height*12+width*pitch/2,0]
p4 = [width*12,height*12,0]
p5 = [width*12,0,0]


entities.add_line(p1,p2)
entities.add_line(p2,p3)
entities.add_line(p3,p4)
entities.add_line(p4,p5)



for i in $a
  entities.add_line(i[0],i[1])
  
  i[1][1] = wcht*12 if i[1][1] == 0
  if i[1][0] <= width*12/2.0
    entities.add_dimension_linear(i[0],i[1],[-20,0,0]).arrow_type = Sketchup::Dimension::ARROW_SLASH
  end 
  if i[1][0] > width*12/2.0
      entities.add_dimension_linear(i[0],i[1],[20,0,0]).arrow_type = Sketchup::Dimension::ARROW_SLASH
  end

end

entities.add_dimension_linear([0,0,0],$a[0][1],[0,-10,0]).arrow_type = Sketchup::Dimension::ARROW_SLASH


$length = width*12
$width = width*12
$bottom = -20

h = width/6/12*12
  c = entities.add_group
  c.entities.add_3d_text('GREINER BUILDING INC.', TextAlignCenter, "Times New Roman", true,false,h*1.5,0,-10,true,1)
  c.material = "black"
  t = Geom::Transformation.new [-$length*0,$bottom-($width/12/2)-h*2,0]
  entities.transform_entities t, c
  # d = entities.add_group
  # d.entities.add_3d_text('2088 250TH ST. WASHINGTON, IA', TextAlignCenter, "Times New Roman", true,false,h,0,-10,true,1)
  # d.material = "white"
  # t = Geom::Transformation.new [-$length*0,$bottom-($width/12/2)-h*3.5,0]
  # entities.transform_entities t, d
  g = entities.add_group
  g.entities.add_3d_text('www.greinerbuildings.com', TextAlignCenter, "Times New Roman", true,false,h,0,-10,true,1)
  g.material = "black"
  t = Geom::Transformation.new [-$length*0,$bottom-($width/12/2)-h*3.5,0]
  entities.transform_entities t, g


  e = entities.add_group
  e.entities.add_3d_text('For: '+worksheet.Cells(182,2).Value + '     '+worksheet.Cells(183,2).Value.strftime("%m/%d/%Y"), TextAlignCenter, "Times New Roman", true,false,h,0,-10,true,1)
  e.material = "black"
  t = Geom::Transformation.new [$length*0.5,$bottom-($width/12/2)-h*2,0]
  entities.transform_entities t,e
  f = entities.add_group
  f.entities.add_3d_text('By: '+worksheet.Cells(184,2).Value + '    (888) 466-4139', TextAlignCenter, "Times New Roman", true,false,h,0,-10,true,1)
  f.material = "black"
  t = Geom::Transformation.new [$length*0.5,$bottom-($width/12/2)-h*3.5,0]
  entities.transform_entities t, f

  g = entities.add_group
  g.entities.add_3d_text('Endwall 2', TextAlignCenter, "Times New Roman", true,false,h*1,0,-10,true,1)
  g.material = "black"
  t = Geom::Transformation.new [$length*0.3,$bottom-($width/12/2)-h*(-1),0]
  entities.transform_entities t, g


Sketchup.send_action "viewTop:"
Sketchup.active_model.active_view.camera.perspective = false

rescue NoMemoryError

ensure
  workbook.Save
  application.Workbooks.Close
  application.quit
end

