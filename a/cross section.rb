require "Sketchup.rb"
require 'win32ole'

entities = Sketchup.active_model.entities

for i in 0..entities.length
  if(entities[0] != nil)
      entities[0].erase!
  end
end

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

begin

width = worksheet.Cells(2,2).Value.to_f
concrete_thick =  worksheet.Cells(505,2).Value.to_f
pitch = worksheet.Cells(5,2).Value.to_f
interior_wall = worksheet.Cells(506,2).Value
out = worksheet.Cells(500,2).Value
out_space = out[out.index('Side Girts, ')+12..out.index('"')-1].to_f
out_n = out[0].to_i
outside_girts = [39]
for i in 1..out_n-1
  outside_girts = outside_girts + [39+i*out_space]
end
roof_steel_length = worksheet.Cells(504,2).Value.to_f
roof = worksheet.Cells(501,2).Value
roof_space = roof[roof.index('per side')+9..roof.index("'")-1].to_f
roof_n = roof[0..roof.index(' ')].to_f

roof_girts = []
for i in 1..roof_n
  roof_girts = roof_girts + [roof_space*i]
end

roof_girts[-1] = roof_steel_length-4 


outside_steel_length = worksheet.Cells(503,2).Value.to_f

overhang = worksheet.Cells(6,2).Value.to_f

height = worksheet.Cells(4,2).Value.to_f - concrete_thick

heel = worksheet.Cells(508,2).Value.to_f
above_ceiling = heel-0.5+3.75-pitch/12.0*overhang-5
above_ceiling = above_ceiling + 5 if overhang == 0

bottom_width = 18

wainscot = worksheet.Cells(10,2).Value.to_f


# concrete above grade

entities.add_edges([0,0,0],[0,concrete_thick,0],[80,concrete_thick,0],[80,0,0],[0,0,0])
entities.add_dimension_linear([80,concrete_thick,0],[80,0,0],[5,0,0]).arrow_type = Sketchup::Dimension::ARROW_NONE
#beneath grade
entities.add_line([0,0,0],[0,-3*12-6,0])
entities.add_line([4,0,0],[4,-3*12-6,0])
entities.add_dimension_linear([0,0,0],[0,-3*12-6,0],[100,0,0]).arrow_type = Sketchup::Dimension::ARROW_NONE
point1 = Geom::Point3d.new(-(bottom_width-4)/2,-3*12-6,0)
point2 = point1 + Geom::Vector3d.new(bottom_width,0,0)
point3 = point2 + Geom::Vector3d.new(0,-6,0)
point4 = point1 + Geom::Vector3d.new(0,-6,0)
entities.add_edges(point1,point2,point3,point4,point1)
entities.add_dimension_linear(point1,point4,[100+(bottom_width-4)/2,0,0]).arrow_type = Sketchup::Dimension::ARROW_NONE
entities.add_dimension_linear(point3,point4,[0,-5,0]).arrow_type = Sketchup::Dimension::ARROW_NONE
#overhang
if overhang > 0
  point1 = Geom::Point3d.new(-1.5,concrete_thick+height+above_ceiling,0)
  entities.add_dimension_linear(point1,point1+Geom::Vector3d.new(-overhang,0,0),[0,-5,0]).arrow_type = Sketchup::Dimension::ARROW_NONE
end
#GE
entities.add_text("GR",[-60,0,0])


def draw_girt(width,height,point1,entities)
  point2 = point1 + Geom::Vector3d.new(0,height,0)
  point3 = point1 + Geom::Vector3d.new(width,height,0)
  point4 = point1 + Geom::Vector3d.new(width,0,0)

  entities.add_edges(point1,point2,point3,point4,point1,point3,point2,point4)
end

draw_girt(1.5,5.5,Geom::Point3d.new(-1.5,0,0),entities)
draw_girt(4,1.5,Geom::Point3d.new(0,concrete_thick,0),entities)

def outside(height,above_ceiling,girts,steel_length,entities,concrete_thick,wainscot)
  entities.add_line([0,height+concrete_thick,0],[120,height+concrete_thick,0])
  entities.add_dimension_linear([70,height+concrete_thick,0],[70,concrete_thick,0],[5,0,0]).arrow_type = Sketchup::Dimension::ARROW_NONE

  entities.add_line([0,0,0],[0,concrete_thick+above_ceiling+height])
  entities.add_line([-2,concrete_thick,0],[-2,concrete_thick+above_ceiling+height])
  entities.add_dimension_linear([-2,concrete_thick+above_ceiling+height,0],[-2,concrete_thick+height,0],[-50,0,0]).arrow_type = Sketchup::Dimension::ARROW_NONE
  entities.add_dimension_linear([-3,concrete_thick+above_ceiling+height-1,0],[-3,concrete_thick+above_ceiling+height-1-steel_length,0],[-60,0,0]).arrow_type = Sketchup::Dimension::ARROW_NONE

  girt_dimension = [[-2,0,0]]
  for i in girts
    draw_girt(1.5,5.5,Geom::Point3d.new(-1.5,i,0),entities)
    girt_dimension = girt_dimension + [[-2,i,0]]
  end
  girt_dimension = girt_dimension + [[-2,concrete_thick+height,0]]
  i = 1
  while i != girt_dimension.length
    entities.add_dimension_linear(girt_dimension[i-1],girt_dimension[i],[-40,0,0]).arrow_type = Sketchup::Dimension::ARROW_NONE
    if i==2
      entities.add_dimension_linear(Geom::Point3d.new(girt_dimension[i-1])+Geom::Vector3d.new(0,5.5,0),girt_dimension[i],[-20,0,0]).arrow_type = Sketchup::Dimension::ARROW_NONE
    end
    i = i + 1
  end

  draw_girt(1.5,3.5,Geom::Point3d.new(-1.5,concrete_thick+height+above_ceiling-3.5,0),entities)

  if wainscot > 1
    point1 = Geom::Point3d.new(-3,concrete_thick+1.5,0)
    entities.add_dimension_linear(point1,point1+Geom::Vector3d.new(0,wainscot,0),[-60,0,0]).arrow_type = Sketchup::Dimension::ARROW_NONE
  end

end


outside(height,above_ceiling,outside_girts,outside_steel_length,entities,concrete_thick,wainscot)

def roof(entities,roof_steel_length,point1,pitch,purlin)
  roof_group = entities.add_group
  entities = roof_group.entities
  entities.add_line(point1, point1+Geom::Vector3d.new(roof_steel_length+0.5,0,0))
  entities.add_dimension_linear(point1, point1+Geom::Vector3d.new(roof_steel_length+0.5,0,0),[0,30,0]).arrow_type = Sketchup::Dimension::ARROW_NONE
  point2 = point1 + Geom::Vector3d.new(0,3.75,0)
  entities.add_line(point2, point2+Geom::Vector3d.new(roof_steel_length,0,0))
  entities.add_dimension_linear(point2, point2+Geom::Vector3d.new(roof_steel_length,3.75,0),[0,40,0]).arrow_type = Sketchup::Dimension::ARROW_NONE

  entities.transform_entities(Geom::Transformation.rotation(point1, [0,0,1], Math::atan(pitch)), roof_group)

  purlin_dimension = [point1]
  for i in purlin
    draw_girt(1.5,3.5,point1+Geom::Vector3d.new(i,0,0),entities)
    purlin_dimension = purlin_dimension + [point1+Geom::Vector3d.new(i,0,0)]
  end

  i = 1
  while i != purlin_dimension.length
    entities.add_dimension_linear(purlin_dimension[i-1],purlin_dimension[i],[0,20,0]).arrow_type = Sketchup::Dimension::ARROW_NONE
    i = i + 1
  end
end

above_ceiling = above_ceiling - 5.5 if overhang == 0
roof_point = Geom::Point3d.new(-overhang-2,concrete_thick+height+above_ceiling+1.2/3.5*pitch,0)
roof(entities,roof_steel_length,roof_point,pitch/12,roof_girts)


truss_point = Geom::Point3d.new(0,concrete_thick+height+2,0)
entities.add_line(truss_point,truss_point + Geom::Vector3d.new(width/24*12.0,pitch*width/24.0,0))
entities.add_line(truss_point + Geom::Vector3d.new(2*12,pitch*2,0),truss_point + Geom::Vector3d.new(2*12,pitch*2,0)+Geom::Vector3d.new(100,0,0))


#heel
point1 = Geom::Point3d.new(0,concrete_thick+height,0)
entities.add_dimension_linear(point1, point1+Geom::Vector3d.new(0,heel,0),[10,0,0]).arrow_type = Sketchup::Dimension::ARROW_NONE

def inside(entities,concrete_thick,height,girts,dimension,inside_steel_length)
  entities.add_line([4,concrete_thick,0],[4,concrete_thick+height,0])
  entities.add_line([6,concrete_thick,0],[6,concrete_thick+height,0])

  for i in girts
    draw_girt(1.5,3.5,Geom::Point3d.new(4,i+concrete_thick,0),entities)
  end

  point1 = Geom::Point3d.new(6,concrete_thick,0)
  for i in dimension
    entities.add_dimension_linear(point1,point1+Geom::Vector3d.new(0,i,0),[20,0,0]).arrow_type = Sketchup::Dimension::ARROW_NONE
    point1 = point1+Geom::Vector3d.new(0,i,0)
  end

  entities.add_dimension_linear(Geom::Point3d.new(6,concrete_thick,0),Geom::Point3d.new(6,concrete_thick,0)+Geom::Vector3d.new(0,inside_steel_length,0),[40,0,0]).arrow_type = Sketchup::Dimension::ARROW_NONE
  entities.add_dimension_linear([6,0,0],point1,[100-6,0,0]).arrow_type = Sketchup::Dimension::ARROW_SLASH
end


 
#inside_girts= [0,48,48*2,48*3,48*4-3.5]
if interior_wall != "no"
  inside_girts = [0]
  inside_dimension = []
  inside_n = (height/12.0/4.0).ceil
  inside_space = height/inside_n
  for i in 1..inside_n
    inside_dimension = inside_dimension + [inside_space]
    inside_girts = inside_girts + [i*inside_space]
  end

  inside_girts[inside_girts.length-1] = inside_girts[inside_girts.length-1]-3.5

  inside(entities,concrete_thick,height,inside_girts,inside_dimension,inside_girts[-1]+2)
else
  entities.add_line([4,concrete_thick,0],[4,concrete_thick+height,0])
end




$length = width
$width = width
$bottom = -50

h = width/6/12
  c = entities.add_group
  c.entities.add_3d_text('GREINER BUILDING INC.', TextAlignCenter, "Times New Roman", true,false,h*1.5,0,-10,true,1)
  c.material = "black"
  t = Geom::Transformation.new [-$length*0.2,$bottom-($width/12/2)-h*2,0]
  entities.transform_entities t, c
  # d = entities.add_group
  # d.entities.add_3d_text('2088 250TH ST. WASHINGTON, IA', TextAlignCenter, "Times New Roman", true,false,h,0,-10,true,1)
  # d.material = "white"
  # t = Geom::Transformation.new [-$length*0.2,$bottom-($width/12/2)-h*3.5,0]
  # entities.transform_entities t, d
  g = entities.add_group
  g.entities.add_3d_text('www.greinerbuildings.com', TextAlignCenter, "Times New Roman", true,false,h,0,-10,true,1)
  g.material = "black"
  t = Geom::Transformation.new [-$length*0.2,$bottom-($width/12/2)-h*3.5,0]
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



  #pitch sign
  pitch_point = Geom::Point3d.new(0,height+8*12,0)
  b = entities.add_dimension_linear(pitch_point,pitch_point + Geom::Vector3d.new(12*5,0,0),[0,0.1,0])
  b.arrow_type = Sketchup::Dimension::ARROW_NONE
  b.text = "12"
  a = entities.add_dimension_linear(pitch_point,pitch_point + Geom::Vector3d.new(0,-pitch*5,0),[-0.1,0,0])
  a.text=pitch.to_s
  a.arrow_type = Sketchup::Dimension::ARROW_NONE

Sketchup.send_action "viewTop:"
Sketchup.active_model.active_view.camera.perspective = false
Sketchup.active_model.active_view.zoom_extents

  rescue NoMemoryError

ensure
  workbook.Save
  application.Workbooks.Close
  application.quit
end