require 'win32ole'
require 'sketchup.rb'


model = Sketchup.active_model
entities = model.active_entities

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
model = Sketchup.active_model
entities = model.entities


colorOptions = []  
  Sketchup.active_model.materials.each { |material|
    colorOptions.push(material.name)
}



$width = worksheet.Cells(2,2).Value.to_f
$length = worksheet.Cells(3,2).Value.to_f
$height = worksheet.Cells(4,2).Value.to_f
$pitch = worksheet.Cells(5,2).Value.to_f
$oh1 = worksheet.Cells(6,2).Value.to_f
$oh2 = worksheet.Cells(7,2).Value.to_f
$oj1 = worksheet.Cells(8,2).Value.to_f
$oj2 = worksheet.Cells(9,2).Value.to_f
$wcht = worksheet.Cells(10,2).Value.to_f+0.01
$apron = worksheet.Cells(11,2).Value.to_f
$heel = worksheet.Cells(12,2).Value.to_f
$wallColor = worksheet.Cells(13,2).Value
$sofitColor = worksheet.Cells(14,2).Value 
$concrete_color = "concrete"
$concrete_height = worksheet.Cells(16,2).Value.to_f
$roofColor = worksheet.Cells(17,2).Value
$faciaColor = worksheet.Cells(18,2).Value
$wainscotColor = worksheet.Cells(19,2).Value

#[15*12, 12*12, 20*12, 6, "SW1", "Kynar Antique Ivory Steel"],[40*12, 12*12, 20*12, 6, "SW1", "Kynar Antique Ivory Steel"]
# overhead_data = []
# $offset_length = 20*12
# $door_height = 12*12
# $door_width = 20*12
# $panel = 6
# $side = "SW1"
# $overheadColor = "Kynar Antique Ivory Steel"


#["EW1", 80, 36, "Kynar Antique Ivory Steel", 50*12, "1","1","right"],
#["SW1", 80, 36, "Kynar Antique Ivory Steel", 5*12, "1","1"]
# walkdoor_data = []
# $walkside = "SW1"
# $walk_height = 80
# $walk_width = 36
# $walk_color = "Kynar Antique Ivory Steel"
# $walk_offset = 10*12
# $walk_window = "1"
# $walk_grid = "1"
# $door_swing = "left"


#cubola
$cupola_number = worksheet.Cells(41,2).Value.to_i
$cub_size = worksheet.Cells(42,2).Value.to_f
$cub_mid = worksheet.Cells(43,2).Value.to_f

#[20*12,14*12,4,36, "Split", "EW1", 12*12, "Kynar Antique Ivory Steel","Kynar Terratone Steel", "blue"]
slide_data = []
$slide_width = 20*12
$slide_height = 14*12
$slide_thickness = 4
$slide_wainscot = 36
$slide_type = "Split"
$slide_side = "EW1"
$slide_offset = 12*12
$slide_color = "Kynar Antique Ivory Steel"
$slide_wainscotcolor = "Kynar Terratone Steel"
$track_color = "blue"

#[-10*12, 30*12, 10*12, "EW1", 2, "Hip"]
#,[10*12, 60*12, 10*12, "EW2", 4, "Hip"]
#porch_data = [[10*12, 60*12, 10*12, "EW2", 0, "Hip"]]
$porch_offset = 0
$porch_width = 0
$porch_length = 0 
$porch_side = ""
$post_number = 5
$porch_type = "Gable"
$porch_height = 8*12
$porch_pitch = 3.5
$porch_overhang = 0



post_data = []
$post_side = ""
$post_offset = 0
$post_width = 0
$post_length = 0

$interior_steel = worksheet.Cells(77,2).Value	


#[14*12,14*12,20*12, "green", "yellow","EW1", 70*12]
hydraulic_data = []
$hydraulic_wainscot = 14*12
$hydraulic_height = 14*12
$hydraulic_width = 20*12
$hydraulic_wainscot_color = "Kynar Beige"
$hydraulic_color = "Kynar Antique Ivory Steel"
$hydraulic_side = "EW1"
$hydraulic_offset = 50*12




$sidelight_side = "SW1"
$sidelight_down = 24

$number_interior_wall = 0
$interior_color = ""
$interior_side = "" #left or right
$interior_distance = 0
$interior_walk = []
$interior_over = []
 

# $lean_height = 12*12
# $lean_length = 10*12
# $lean_width = 100*12
# $lean_pitch = 3.5
# $lean_wall = [1,1,0] #left, mid,right
# $lean_side = "SW1"
# $lean_offset = 12*12
 
#["EW1",2*12, 5*12, 5*12,4*12, "Slider"]
window_data = []
	$window_side = "EW1"
	$window_offset = 2*12
	$window_height = 5*12
	$window_length = 5*12
	$window_width = 4*12
	$window_type = "Slider"
	$window_color = "yellow"

#[22*12,30*12, "SW1", 14*12]
concrete_data = []
$concrete_width = 22*12
$concrete_length = 40*12
$concrete_side = "EW1"
$concrete_offset = 14*12



$grass_color = "green"

$gambrel_height = worksheet.Cells(122,2).Value.to_f
$brian = worksheet.Cells(123,2).Value.to_f


$wainscot_corner = worksheet.Cells(156,2).Value
$main_color = worksheet.Cells(157,2).Value

$ridge = worksheet.Cells(159,2).Value

$facia = 5.5
$rake = worksheet.Cells(160,2).Value

$slide_frameColor = worksheet.Cells(171,2).Value
$GBcolor = "black"#worksheet.Cells(169,2).Value

$top = worksheet.Cells(174,2).Value.to_f
$bottom = worksheet.Cells(175,2).Value.to_f
$left = worksheet.Cells(176,2).Value.to_f
$right = worksheet.Cells(177,2).Value.to_f
$north = worksheet.Cells(179,2).Value

$porch_post_color = worksheet.Cells(164,2).Value
$porch_header_color = worksheet.Cells(165,2).Value
$porch_ceiling_color = worksheet.Cells(166,2).Value
$porch_roof_color = worksheet.Cells(167,2).Value
$wcht = $wcht + $concrete_height
$gableht1 = (((0.5 * $width)+$oh1)*$pitch)/12
$gableht2 = (((0.5 * $width)+$oh2)*$pitch)/12
$posX = 0
$posY = 0
$posZ = 0



    $facia = 5.5
    $corner = 5.5


  
$height1 = $height
$height2 = $height
  
    if ($oh1 > 0) 
      $height1 = $height1+$heel-($pitch*$oh1/12)-3.75-2
    else
      $height1 = $height1 + $heel + 1.25-7
    end
    
	if ($oh2 > 0) 
      $height2 = $height2+$heel-($pitch*$oh2/12)-3.75-2
    else
      $height2 = $height2 + $heel + 1.25-7
    end
	
$height = [$height1,$height2].max

  
  def get_heel
    if $width < 40
      $heel=6
    elsif $width <50
      $heel=8
    elsif $width <60
      $heel=10
    elsif $width <66
      $heel=12
    elsif $width <80
      $heel=14
    elsif $width <100
      $heel=16
    else
      $heel=18
    end

    return $heel  
  end
$heel = get_heel()

def get_ew2_points(posX, posY, posZ)
    pts = []
    posX = posX + $length

   
      pts.push([posX, posY, posZ])
      pts.push([posX, (posY + $width), posZ])
	  pts.push([posX, (posY + $width), (posZ + $height2)])
	if ($oh2 > 0)
	  pts.push([posX, (posY + $width + $oh2), (posZ + $height2)])
	end
    pts.push([posX, (posY + ($width/2)), (posZ + $height2 + $gableht2)])
    
	if($oh1 > 0)
	  pts.push([posX, (posY-$oh1), (posZ + $height1)])
      
    end
	pts.push([posX, posY, (posZ + $height1)])
    return pts
  end

  def get_facia_ew1_points(posX, posY, posZ)
    posZ = posZ + $height1
    posY = posY - $oh1

    pts = []
    pts[0] = [posX - $oj1, posY, posZ]
    pts[1] = [posX + $oj2 + $length, posY, posZ]
    pts[2] = [posX + $oj2 + $length, posY, posZ + $facia]
    pts[3] = [posX - $oj1, posY, posZ + $facia]

    return pts
  end

  def get_facia_ew2_points(posX, posY, posZ)
    posZ = posZ + $height2
    posY = posY + $oh2 + $width

    pts = []
    pts[0] = [posX - $oj1, posY, posZ]
    pts[1] = [posX + $oj2 + $length, posY, posZ]
    pts[2] = [posX + $oj2 + $length, posY, posZ + $facia]
    pts[3] = [posX - $oj1, posY, posZ + $facia] 

    return pts
  end

  def get_oh_sofit_sw1_points(posX, posY, posZ)
    posZ = posZ + $height1

    pts = []
    pts[0] = [posX, posY, posZ]
    pts[1] = [posX, (posY - $oh1), posZ]
    pts[2] = [posX + $length, (posY - $oh1), (posZ)]
    pts[3] = [posX + $length, posY, posZ]

    return pts
  end

  def get_oh_sofit_sw2_points(posX, posY, posZ)
    posY = posY + $width
    posZ = posZ + $height2

    pts = []
    pts[0] = [posX, posY, posZ]
    pts[1] = [posX, (posY + $oh2), posZ]
    pts[2] = [posX + $length, (posY + $oh2), (posZ)]
    pts[3] = [posX + $length, posY, posZ]

    return pts
  end

  def get_oj_sofit_ew1_right_points(posX, posY, posZ)
    posZ = posZ + $height1
    
    pts = []
    pts[0] = [posX, posY - $oh1, posZ]
    pts[1] = [posX - $oj1, (posY - $oh1), posZ]
    pts[2] = [posX - $oj1, posY + ($width / 2), posZ + $gableht1]
    pts[3] = [posX, posY + ($width / 2), posZ + $gableht1]

    return pts
  end

  def get_oj_sofit_ew1_left_points(posX, posY, posZ)
    posZ = posZ + $height2

    pts = []
    pts[0] = [posX, posY + $width + $oh2, posZ]
    pts[1] = [posX - $oj1, posY + $width + $oh2, posZ]
    pts[2] = [posX - $oj1, posY + ($width / 2), posZ + $gableht2]
    pts[3] = [posX, posY + ($width / 2), posZ + $gableht2]

    return pts
  end

  def get_oj_sofit_ew2_right_points(posX, posY, posZ) 
    posZ = posZ + $height2

    pts = []
    pts[0] = [posX + $length, posY + $width + $oh2, posZ]
    pts[1] = [posX + $length + $oj2, posY + $width + $oh2, posZ]
    pts[2] = [posX + $length + $oj2, posY + ($width / 2), posZ + $gableht2]
    pts[3] = [posX + $length, posY + ($width / 2), posZ + $gableht2]

    return pts
  end

  def get_oj_sofit_ew2_left_points(posX, posY, posZ) 
    posZ = posZ + $height1

    pts = []
    pts[0] = [posX + $length, posY - $oh1, posZ]
    pts[1] = [posX + $length + $oj2, (posY-$oh1), posZ]
    pts[2] = [posX + $length + $oj2, posY + ($width / 2), posZ + $gableht1]
    pts[3] = [posX + $length, posY + ($width / 2), posZ + $gableht1]

    return pts
  end

  
  
  def get_sw1_points(posX, posY, posZ)
    # no OHs. just add facia to building.height
    pts = []
    pts[0] = [posX, posY, posZ]
    pts[3] = [(posX + $length), posY, posZ]
    #   add facia ht
    pts[1] = [posX, posY, (posZ + $height1)]
    #   gable peak
    pts[2] = [(posX + $length), posY, (posZ + $height1)]
    
    return pts
  end

  def get_sw2_points(posX, posY, posZ)
    posY = posY + $width

    pts = []
    pts[0] = [posX, posY, posZ]
    pts[1] = [posX, posY, (posZ + $height2)]
    pts[2] = [(posX + $length), posY, (posZ + $height2)]
    pts[3] = [(posX + $length), posY, posZ]

    return pts
  end

  def get_facia_ew1_right_points(posX, posY, posZ)
    posZ = posZ + $height1
    pts = []
    
    pts[0] = [posX - $oj1, posY - $oh1, posZ]
    pts[1] = [posX - $oj1, posY - $oh1, posZ + $facia]
    pts[2] = [posX - $oj1, posY + ($width / 2), posZ + $gableht1 + $facia]
    pts[3] = [posX - $oj1, posY + ($width / 2), posZ + $gableht1]

    return pts
  end

  def get_facia_ew1_left_points(posX, posY, posZ)
    posZ = posZ + $height2
    pts = []

    pts[0] = [posX - $oj1, posY + $oh2 + $width, posZ]
    pts[1] = [posX - $oj1, posY + $oh2 + $width, posZ + $facia]
    pts[2] = [posX - $oj1, posY + ($width / 2), posZ + $gableht2 + $facia]
    pts[3] = [posX - $oj1, posY + ($width / 2), posZ + $gableht2]

    return pts
  end

  def get_facia_ew2_right_points(posX, posY, posZ)
    posX = posX + $length + $oj2
    posZ = posZ + $height1
    pts = []

    pts[0] = [posX, posY - $oh1, posZ]
    pts[1] = [posX, posY - $oh1, posZ + $facia]
    pts[2] = [posX, posY + ($width / 2), posZ + $gableht1 + $facia]
    pts[3] = [posX, posY + ($width / 2), posZ + $gableht1]

    return pts
  end

  def get_facia_ew2_left_points(posX, posY, posZ)
    posX = posX + $length + $oj2
    posZ = posZ + $height2
    pts = []

    pts[0] = [posX, posY + $oh2 + $width, posZ]
    pts[1] = [posX, posY + $oh2 + $width, posZ + $facia]
    pts[2] = [posX, posY + ($width / 2), posZ + $gableht2 + $facia]
    pts[3] = [posX, posY + ($width / 2), posZ + $gableht2]

    return pts    
  end

  def get_roof_sw1_points(posX, posY, posZ)
    posZ = posZ + $height1 + $facia
    # no OHs. just add facia to $height
    pts = []
    pts[0] = [posX - $oj1, posY - $oh1, posZ]
    pts[1] = [posX - $oj1, posY + ($width / 2), (posZ + $gableht1)]
    pts[2] = [posX + $length + $oj2, posY + ($width / 2), (posZ + $gableht1)]
    pts[3] = [posX + $length + $oj2, posY - $oh1, posZ]

    return pts
  end

  def get_roof_sw2_points(posX, posY, posZ)
    posZ = posZ + $height2 + $facia
    posY = posY + $width

    # no OHs. just add facia to $height
    pts = []
    pts[0] = [posX - $oj1, posY + $oh2, posZ]
    pts[1] = [posX - $oj1, posY - ($width / 2), (posZ + $gableht2)]
    pts[2] = [posX + $length + $oj2, posY - ($width / 2), (posZ + $gableht2)]
    pts[3] = [posX + $length + $oj2, posY + $oh2, posZ]	 

    return pts
  end

  def get_wainscot_sw1_points(posX, posY, posZ)
    pts = []

    pts[0] = [posX + $corner, posY, posZ]
    pts[1] = [posX + $corner, posY, (posZ + $wcht)]
    pts[2] = [(posX + $length) - $corner, posY, (posZ + $wcht)]
    pts[3] = [(posX + $length) - $corner, posY, posZ]

    return pts
  end

  def get_wainscot_sw2_points(posX, posY, posZ)
    pts = []

    pts[0] = [posX + $corner, posY + $width, posZ]
    pts[1] = [posX + $corner, posY + $width, (posZ + $wcht)]
    pts[2] = [(posX + $length) - $corner, posY + $width, (posZ + $wcht)]
    pts[3] = [(posX + $length) - $corner, posY + $width, posZ]

    return pts
  end

  def get_wainscot_ew1_points(posX, posY, posZ)
    pts = []

    pts[0] = [posX, posY + $corner, posZ]
    pts[1] = [posX, posY + $corner, (posZ + $wcht)]
    pts[2] = [posX, posY + $width - $corner, (posZ + $wcht)]
    pts[3] = [posX, posY + $width - $corner, posZ]

    return pts
  end


    
	
	
  def get_wainscot_ew2_points(posX, posY, posZ)
    pts = []

    pts[0] = [posX + $length, posY + $corner, posZ]
    pts[1] = [posX + $length, posY + $corner, (posZ + $wcht)]
    pts[2] = [posX + $length, posY + $width - $corner, (posZ + $wcht)]
    pts[3] = [posX + $length, posY + $width - $corner, posZ]

    return pts
  end

  def has_overhang1?
    return $oh1 > 0
  end
  
  def has_overhang2?
    return $oh2 > 0
  end

  def has_overjet1?
    return $oj1 > 0
  end
  
  def has_overjet2?
	return $oj2 > 0
  end

  def has_wainscot?
    return $wcht > 0
  end

  def has_apron?
    return $apron > 0
  end
  



 #build a gable roof porch 
def build_gr(entities,r,t)

gr_group = entities.add_group

origin = Geom::Point3d.new(0,0,0)
v1 = Geom::Vector3d.new($gr_width/2,0,0)
v2 = Geom::Vector3d.new(0,0,$gr_width*$gr_pitch/12)
v3 = Geom::Vector3d.new(0, $gr_length,0)

f1 = gr_group.entities.add_face(origin,origin-v3,origin-v3+v2+v1,origin+v2+v1)
f1.material = $roofColor
f2 = gr_group.entities.add_face(origin-v3+v2+v1,origin+v2+v1, origin+v1+v1, origin+v1+v1-v3)
f2.material = $roofColor

if($gr_wall[1] == 2)
	f5 = gr_group.entities.add_face([0,-$gr_length,0],[6, -$gr_length,0],[6, 6-$gr_length,0],[0,6-$gr_length,0])
	f5.pushpull $gr_height
	f6 = gr_group.entities.add_face([$gr_width,-$gr_length,0],[$gr_width-6, -$gr_length,0],[$gr_width-6, 6-$gr_length,0],[$gr_width,6-$gr_length,0])
	f6.pushpull $gr_height
end

if($gr_wall[0] == 1)
	f3 = gr_group.entities.add_face([0,0,0],[0,0,-$gr_height],[0,-$gr_length, -$gr_height],[0, -$gr_length,0])
	f3.material = $wallColor
end

if($gr_wall[2] == 1)
	f3 = gr_group.entities.add_face([$gr_width,0,0],[$gr_width,0,-$gr_height],[$gr_width,-$gr_length, -$gr_height],[$gr_width, -$gr_length,0])
	f3.material = $wallColor
end

entities.transform_entities r, gr_group
entities.transform_entities t, gr_group

end

# buid and then move gable roof porch 
def create_gr(entities)
	if($gr_side == "SW1")
		r = Geom::Transformation.rotation [0, 0, 0], [0,0,1], 360.degrees
		t = Geom::Transformation.new [$gr_offset,0,$gr_height]
		build_gr(entities,r,t)
	end
	if($gr_side == "SW2")
		r = Geom::Transformation.rotation [0, 0, 0], [0,0,1], 180.degrees
		t = Geom::Transformation.new [$length-$gr_offset,$width,$gr_height]
		build_gr(entities,r,t)
	end
	if($gr_side == "EW1")
		r = Geom::Transformation.rotation [0, 0, 0], [0,0,1], 270.degrees
		t = Geom::Transformation.new [0,$width-$gr_offset,$gr_height]
		build_gr(entities,r,t)
	end
	if($gr_side == "EW2")
		r = Geom::Transformation.rotation [0, 0, 0], [0,0,1], 90.degrees
		t = Geom::Transformation.new [$length,$gr_offset,$gr_height]
		build_gr(entities,r,t)
	end
	
end

########create_gr here


#build lean to  
def build_lean(entities,r,t)



  lean_group = entities.add_group

  lean_group.entities.add_dimension_linear([0,0,0],[0,$lean_length,0], [-30,0,0]).arrow_type = Sketchup::Dimension::ARROW_SLASH
  lean_group.entities.add_dimension_linear([0,0,0],[$lean_width,0,0], [0,-60,0]).arrow_type = Sketchup::Dimension::ARROW_SLASH
  lean_group.entities.add_text("Lean-to", [$lean_width/2,$lean_length/2,-2])


  o = Geom::Point3d.new(-$lean_overjet,-$lean_overhang,$lean_height-$lean_overhang*$lean_pitch/12)
  h = Geom::Vector3d.new(0,$lean_length+$lean_overhang, ($lean_length+$lean_overhang)*$lean_pitch/12)
  f = Geom::Vector3d.new(0,0,$facia)
  v = Geom::Vector3d.new($lean_width+2*$lean_overjet,0,0)
  p1 = o
  p2 = o + v
  p3 = o + v + h
  p4 = o + h
  p5 = o + f
  p6 = o + f + v
  p7 = o + f + v + h
  p8 = o + f + h

  lean_group.entities.add_face(p5,p6,p7,p8).material = $roofColor
  lean_group.entities.add_face(p1,p4,p8,p5).material = $faciaColor
  lean_group.entities.add_face(p2,p3,p7,p6).material = $faciaColor
  lean_group.entities.add_face(p1,p2,p6,p5).material = $faciaColor
  lean_group.entities.add_face(p1,p2,p3,p4)

  lean_group.entities.add_cline([-$lean_overjet,-$lean_overhang,0],[$lean_width+$lean_overjet,-$lean_overhang,0])
  lean_group.entities.add_cline([$lean_width+$lean_overjet,-$lean_overhang,0],[$lean_width+$lean_overjet,$lean_length,0])
  lean_group.entities.add_cline([-$lean_overjet,-$lean_overhang,0],[-$lean_overjet,$lean_length,0])


  if($lean_wall[0] == 1)
    pl1 = [$lean_width, 0, 0]
    pl2 = [$lean_width, $lean_length,0]
    pl3 = [$lean_width,$lean_length,$lean_height+$lean_length*$lean_pitch/12]
    pl4 = [$lean_width,0,$lean_height]
    lean_group.entities.add_face(pl1,pl2,pl3,pl4).material = $wallColor
  end

  if($lean_wall[1] == 1)
    pm1 = [$lean_width,0,0]
    pm2 = [0,0,0]
    pm3 = [0,0,$lean_height]
    pm4 = [$lean_width,0,$lean_height]
    a = lean_group.entities.add_face(pm1,pm2,pm3,pm4)
    a.back_material = $wallColor
    a.material = $wallColor
  end


  p = 6
  if($lean_wall[1] >= 2)
    c = 7
    #$lean_width = $lean_width-2*$lean_overhang
    lean_group.entities.add_face([0,$lean_length,0],[0,0,0],[$lean_width,0,0],[$lean_width,$lean_length,0],[$lean_width-c,$lean_length,0],[$lean_width-c,c,0],[c,c,0],[c,$lean_length,0],[0,$lean_length,0]).pushpull -3
    lean_group.entities.add_line([0,$lean_length,0],[c,$lean_length,0]).faces[0].erase!

    n = $lean_wall[1] 
    # entities.add_face([0,0,0,],[0,p,0],[p,p,0],[p,0,0]).pushpull -$lean_height-$facia
    # entities.add_face([$lean_width,0,0],[$lean_width,p,0],[$lean_width-p,p,0],[$lean_width-p,0,0])
    d = ($lean_width - p*n)/(n-2+1)
    for i in 0..n-1 
      a = lean_group.entities.add_face([i*(p+d),0,0,],[i*(p+d),p,0],[i*(p+d)+p,p,0],[i*(p+d)+p,0,0])
    end
    d = $lean_width/(n-1)
    if n>2
        for i in 0..n-2
            lean_group.entities.add_dimension_linear([i*d,0,0],[(i+1)*d,0,0], [0,-30,0]).arrow_type = Sketchup::Dimension::ARROW_SLASH
        end
    end

    



  end

  if($lean_wall[2] == 1)
    pr1 = [0,0,0]
    pr2 = [0,$lean_length,0]
    pr3 = [0,$lean_length,$lean_height+$lean_length*$lean_pitch/12]
    pr4 = [0,0,$lean_height]
    lean_group.entities.add_face(pr1,pr2,pr3,pr4).back_material = $wallColor
  end

  entities.transform_entities r, lean_group
  entities.transform_entities t, lean_group
  
end

	

 


# build and move lean to
def create_lean(entities)
	if($lean_height > 0)
		if($lean_side == "SW1")
			r = Geom::Transformation.rotation [0,0,0],[0,0,1],360.degrees
			t = Geom::Transformation.new [$lean_offset,-$lean_length,0]
			build_lean(entities, r, t)
		end
		if($lean_side == "SW2")
			r = Geom::Transformation.rotation [0,0,0],[0,0,1],180.degrees
			t = Geom::Transformation.new [$length-$lean_offset,+$width+$lean_length,0]
			build_lean(entities, r, t)
		end
		if($lean_side == "EW1")
			r = Geom::Transformation.rotation [0,0,0],[0,0,1],270.degrees
			t = Geom::Transformation.new [-$lean_length,+$width-$lean_offset,0]
			build_lean(entities, r, t)
		end
		if($lean_side == "EW2")
			r = Geom::Transformation.rotation [0,0,0],[0,0,1],90.degrees
			t = Geom::Transformation.new [$length+$lean_length,$lean_offset,0]
			build_lean(entities, r, t)
		end
	end
end
$size1 = 12
if($size1>0)
	for i in 1..$size1
		
		$lean_length = worksheet.Cells(97,1+i).Value.to_f
		$lean_width = worksheet.Cells(98,1+i).Value.to_f
		$lean_pitch = worksheet.Cells(99,1+i).Value.to_f
		$lean_wall = [worksheet.Cells(101,1+i).Value.to_f,worksheet.Cells(102,1+i).Value.to_f,worksheet.Cells(103,1+i).Value.to_f] #left, mid,right
		$lean_side = worksheet.Cells(104,1+i).Value
		$lean_offset = worksheet.Cells(105,1+i).Value.to_f
		$lean_height = $height-worksheet.Cells(96,1+i).Value.to_f-$lean_length*$lean_pitch/12
    $lean_overhang = worksheet.Cells(134,1+i).Value.to_f
    $lean_overjet = worksheet.Cells(135,1+i).Value.to_f
		create_lean(entities)


    #$lean_wall[1] = 1
	end
end


$size1 = 2
if($size1>0)
	for i in 1..$size1
		$gr_length = worksheet.Cells(138,1+i).Value.to_f
		$gr_width = worksheet.Cells(139,1+i).Value.to_f
		$gr_pitch = worksheet.Cells(140,1+i).Value.to_f
		$gr_height = worksheet.Cells(141,1+i).Value.to_f
		$gr_wall = [worksheet.Cells(145,1+i).Value.to_f,worksheet.Cells(146,1+i).Value.to_f,worksheet.Cells(147,1+i).Value.to_f]
		$gr_offset = worksheet.Cells(142,1+i).Value.to_f
		$gr_side = worksheet.Cells(143,1+i).Value
	
		create_gr(entities)
	end
end

#build walkdoor
def create_walkdoor(entities)
  a = 4
  trim_width = -7
  if($walkside)
	if($walkside == "EW1")
		#new_face = entities.add_face([a, width - offset_length, 0],[9, width - offset_length, door_height],[9, width - offset_length - door_length, door_height],[9, width - offset_length - door_length, 0])
		door = entities.add_face([a, $width - $walk_offset, 0],[a, $width - $walk_offset, $walk_height],[a, $width - $walk_offset - $walk_width, $walk_height],[a, $width - $walk_offset - $walk_width, 0])
 		door.pushpull -a, true
		door.material = $walk_color 
		
		#door = entities.add_face([a, $width - $walk_offset - 10, $walk_height/2 + 10],[a, $width - $walk_offset - 10, $walk_height -10],[a, $width - $walk_offset - $walk_width + 10, $walk_height - 10],[a, $width - $walk_offset - $walk_width - 10, $walk_height/2 + 10])
		
		if($walk_window == 1)
			window = entities.add_face([a, $width - $walk_offset+trim_width, $walk_height/2 -3], [a, $width - $walk_offset + trim_width, $walk_height + trim_width], [a, $width - $walk_offset - $walk_width - trim_width, $walk_height + trim_width], [a, $width - $walk_offset - $walk_width - trim_width, $walk_height/2 - 3])
			window.material = "[Translucent_Glass_Blue]"
			window.back_material = window.material
			if($walk_grid == 1)
				# origin = Geom::Point3d.new(a, $width - $walk_offset - trim_width, $walk_height)
				# v1 = Geom::Vector3d.new(0, $wa/3,0)
				# v2 = Geom::Vector3d.new(0,0,$walk_height + trim_width + 3 - $walk_height/2)
				
				# entities.add_line origin+v1, origin+v1+v2+v2+v2
				for i in 1..2
					entities.add_line([a, $width - $walk_offset + trim_width-((($walk_width+trim_width*2)/3)*i), $walk_height/2 -3], [a, $width - $walk_offset + trim_width-((($walk_width+trim_width*2)/3)*i), $walk_height/2 -3 +($walk_height+trim_width+3-$walk_height/2)])
					entities.add_line([a, $width - $walk_offset+trim_width, $walk_height/2 -3 + (($walk_height+trim_width+3-$walk_height/2)/3)*i],[a, $width - $walk_offset+trim_width - ($walk_width+trim_width*2), $walk_height/2 -3 + (($walk_height+trim_width+3-$walk_height/2)/3)*i])
				end
			end
		end
		
		if($door_swing == "left")
			center_point = Geom::Point3d.new(a-2,$width-$walk_offset-4,$walk_height/2 -2)
			edges = entities.add_arc([a, $width-$walk_offset, -3], Geom::Vector3d.new(1,0,0), Geom::Vector3d.new(0,0,1), $walk_width, 270.degrees, 300.degrees)
			entities.add_line([a, $width-$walk_offset, -3],edges[1].end)
		else
			center_point = Geom::Point3d.new(a-2,$width-$walk_offset-$walk_width+4,$walk_height/2 -2)
			edges = entities.add_arc([a, $width-$walk_offset-$walk_width, -3], Geom::Vector3d.new(1,0,0), Geom::Vector3d.new(0,0,1), $walk_width, 60.degrees, 90.degrees)
			entities.add_line([a, $width-$walk_offset-$walk_width, -3],edges[0].start)
		end
		radius = 2
		normal_vector = Geom::Vector3d.new(1,0,0)
		edgearray = entities.add_circle center_point, normal_vector, radius
		edgearray[0].find_faces
		face = edgearray[0].faces[0]
		face.material = "gray"
		face.pushpull 4
		
		if($wcht>0)
			l =	entities.add_line([0, $width - $walk_offset - $walk_width, $wcht],[0, $width - $walk_offset, $wcht])
			l.faces[0].erase!
			l.faces[0].erase!
			l.erase!
		end
		
		if($concrete_height>0)
			entities.add_line([0, $width - $walk_offset - $walk_width, $concrete_height],[0, $width - $walk_offset, $concrete_height]).erase!
		end
		
		# new_face.pushpull 0.001
		# entities.add_face([0, $width-$walk_offset, 0], [a, $width-$walk_offset, 0], [a, $width - $walk_offset, $walk_height], [0, $width- $walk_offset, $walk_height]).reverse!.material = $faciaColor
		# entities.add_face([0, $width- $walk_offset, $walk_height], [a, $width- $walk_offset, $walk_height], [a, $width- $walk_offset - $walk_width, $walk_height],[0, $width- $walk_offset - $walk_width, $walk_height]).reverse!.material = $faciaColor
		# entities.add_face([a, $width- $walk_offset - $walk_width, $walk_height],[0, $width- $walk_offset - $walk_width, $walk_height],[0, $width- $walk_offset - $walk_width, 0],[a, $width- $walk_offset - $walk_width, 0]).reverse!.material = $faciaColor
		
		
		
		
	end
	if($walkside == "EW2")
		#new_face = entities.add_face([$length-9, , 0],[$length-9, width - walk_offset, walk_height],[$length-9, width - walk_offset - door_length, walk_height],[9, width - walk_offset - door_length, 0])
		door = entities.add_face([$length-a, $walk_width + $walk_offset, 0],[$length-a, $walk_width + $walk_offset, $walk_height],[$length-a, $walk_offset, $walk_height],[$length-a, $walk_offset, 0])
		door.material = $walk_color
		door.back_material = $walk_color
		 		door.pushpull +a, true
		if($walk_window == 1)
		window = entities.add_face([$length-a, $walk_offset-trim_width, $walk_height/2 - 3 ], [$length-a, $walk_offset - trim_width, $walk_height + trim_width], [$length-a, $walk_offset + $walk_width + trim_width, $walk_height + trim_width], [$length-a, $walk_offset + $walk_width + trim_width, $walk_height/2 - 3])
		window.material = "[Translucent_Glass_Blue]"
		window.back_material = window.material
		if($walk_grid == 1)
				for i in 1..2
					entities.add_line([$length-a, $walk_offset-trim_width + ((($walk_width+trim_width*2)/3)*i), $walk_height/2 - 3 ],[$length-a, $walk_offset-trim_width + ((($walk_width+trim_width*2)/3)*i), $walk_height/2 - 3 + ($walk_height+trim_width+3-$walk_height/2) ])
					entities.add_line([$length-a, $walk_offset-trim_width , $walk_height/2 - 3 + (($walk_height+trim_width+3-$walk_height/2)/3)*i],[$length-a, $walk_offset-trim_width + ($walk_width+trim_width*2), $walk_height/2 - 3  + (($walk_height+trim_width+3-$walk_height/2)/3)*i])
				end
			end
		end
		
		
		
		#center_point = Geom::Point3d.new($length - a + 2,$walk_offset+4,$walk_height/2 -2)
		if($door_swing == "left")
			center_point = Geom::Point3d.new($length - a + 2,$walk_offset+4,$walk_height/2 -2)
			edges = entities.add_arc([$length-a, $walk_offset, -3], Geom::Vector3d.new(1,0,0), Geom::Vector3d.new(0,0,1), $walk_width, 90.degrees, 120.degrees)
			entities.add_line([$length-a, $walk_offset, -3],edges[1].end)
		else
			center_point = Geom::Point3d.new($length - a + 2,$walk_offset+$walk_width-4,$walk_height/2 -2)
			edges = entities.add_arc([$length-a, $walk_offset+$walk_width, -3], Geom::Vector3d.new(1,0,0), Geom::Vector3d.new(0,0,1), $walk_width, 240.degrees, 270.degrees)
			entities.add_line([$length-a, $walk_offset+$walk_width, -3],edges[0].start)
		end
		
		radius = 2
		normal_vector = Geom::Vector3d.new(1,0,0)
		edgearray = entities.add_circle center_point, normal_vector, radius
		edgearray[0].find_faces
		face = edgearray[0].faces[0]
				face.material = "gray"
		face.pushpull -4
		
		if($wcht>0)
			l =	entities.add_line([$length, $walk_offset, $wcht],[$length, $walk_offset + $walk_width, $wcht])
			l.faces[0].erase!
			l.faces[0].erase!
			l.erase!
		end
		if($concrete_height>0)
			entities.add_line([$length, $walk_offset, $concrete_height],[$length, $walk_offset + $walk_width, $concrete_height]).erase!
		end
	end
	if($walkside == "SW1")
		
		door = entities.add_face([$walk_offset, a, 0],[$walk_offset, a, $walk_height],[$walk_offset + $walk_width, a, $walk_height],[$walk_offset + $walk_width, a, 0])
 		door.pushpull -a, true
		door.material = $walk_color
		if($walk_window == 1)
		window = entities.add_face([$walk_offset - trim_width, a,  $walk_height/2 - 3], [$walk_offset - trim_width, a, $walk_height + trim_width], [$walk_offset + $walk_width + trim_width, a, $walk_height + trim_width], [$walk_offset + $walk_width + trim_width, a,  $walk_height/2 - 3])
		window.material = "[Translucent_Glass_Blue]"
		window.back_material = window.material
		if($walk_grid == 1)
				for i in 1..2
					entities.add_line([$walk_offset - trim_width+((($walk_width+trim_width*2)/3)*i), a,  $walk_height/2 - 3],[$walk_offset - trim_width + ((($walk_width+trim_width*2)/3)*i), a,  $walk_height/2 - 3 + ($walk_height+trim_width+3-$walk_height/2)])
					entities.add_line([$walk_offset - trim_width, a,  $walk_height/2 - 3 + (($walk_height+trim_width+3-$walk_height/2)/3)*i],[$walk_offset - trim_width+($walk_width+trim_width*2), a,  $walk_height/2 - 3+(($walk_height+trim_width+3-$walk_height/2)/3)*i])
				end
			end
		end
		
		if($door_swing == "left")
			center_point = Geom::Point3d.new($walk_offset+4, a - 2, $walk_height/2 -2)
			edges = entities.add_arc([$walk_offset, a, -3], Geom::Vector3d.new(1,0,0), Geom::Vector3d.new(0,0,1), $walk_width, 0.degrees, 30.degrees)
			entities.add_line([$walk_offset, a, -3],edges[1].end)
		else
			center_point = Geom::Point3d.new($walk_offset+$walk_width-4, a - 2, $walk_height/2 -2)
			edges = entities.add_arc([$walk_offset+$walk_width, a, -3], Geom::Vector3d.new(1,0,0), Geom::Vector3d.new(0,0,1), $walk_width, 150.degrees, 180.degrees)
			entities.add_line([$walk_offset+$walk_width, a, -3],edges[0].start)
		end
		radius = 2
		normal_vector = Geom::Vector3d.new(0,1,0)
		edgearray = entities.add_circle center_point, normal_vector, radius
		edgearray[0].find_faces
		face = edgearray[0].faces[0]
				face.material = "gray"
		face.pushpull 4
		
		if($wcht>0)
			l =	entities.add_line([$walk_offset, 0, $wcht],[$walk_offset + $walk_width, 0, $wcht])
			l.faces[0].erase!
			l.faces[0].erase!
			l.erase!
		end
		
		if($concrete_height>0)
			entities.add_line([$walk_offset, 0, $concrete_height],[$walk_offset + $walk_width, 0, $concrete_height]).erase!
		end
	end
	if($walkside == "SW2")
		door = entities.add_face([$length - $walk_offset, $width - a, 0],[$length - $walk_offset, $width - a, $walk_height],[$length - $walk_offset - $walk_width, $width - a, $walk_height],[$length - $walk_offset - $walk_width, $width - a, 0])
 		door.pushpull -a, true
		door.material = $walk_color
		if($walk_window == 1)
		window = entities.add_face([$length - $walk_offset + trim_width, $width-a, $walk_height/2 - 3], [$length - $walk_offset + trim_width, $width-a, $walk_height + trim_width], [$length - $walk_offset - trim_width - $walk_width, $width-a, $walk_height + trim_width], [$length-($walk_offset + $walk_width + trim_width), $width-a, $walk_height/2 - 3])
		window.material = "[Translucent_Glass_Blue]"
		window.back_material = window.material
		if($walk_grid == 1)
				for i in 1..2
					entities.add_line([$length - $walk_offset + trim_width - ((($walk_width+trim_width*2)/3)*i), $width-a, $walk_height/2 - 3],[$length - $walk_offset + trim_width - ((($walk_width+trim_width*2)/3)*i), $width-a, $walk_height/2 - 3 + ($walk_height+trim_width+3-$walk_height/2)])
					entities.add_line([$length - $walk_offset + trim_width, $width-a, $walk_height/2 - 3 + (($walk_height+trim_width+3-$walk_height/2)/3)*i],[$length - $walk_offset + trim_width - ($walk_width+trim_width*2), $width-a, $walk_height/2 - 3 + (($walk_height+trim_width+3-$walk_height/2)/3)*i])
				end
			end
		end
		
		if($door_swing == "left")
			center_point = Geom::Point3d.new($length - $walk_offset - 4, $width - a + 2, $walk_height/2 -2)
			edges = entities.add_arc([$length-$walk_offset, $width-a, -3], Geom::Vector3d.new(1,0,0), Geom::Vector3d.new(0,0,1), $walk_width, 180.degrees, 210.degrees)
			entities.add_line([$length-$walk_offset, $width-a, -3],edges[1].end)
		else
			center_point = Geom::Point3d.new($length - $walk_offset - $walk_width+4, $width - a + 2, $walk_height/2 -2)
			edges = entities.add_arc([$length-$walk_offset-$walk_width, $width-a, -3], Geom::Vector3d.new(1,0,0), Geom::Vector3d.new(0,0,1), $walk_width, 330.degrees, 360.degrees)
			entities.add_line([$length-$walk_offset-$walk_width, $width-a, -3],edges[0].start)
		end
		radius = 2
		normal_vector = Geom::Vector3d.new(0,1,0)
		edgearray = entities.add_circle center_point, normal_vector, radius
		edgearray[0].find_faces
		face = edgearray[0].faces[0]
				face.material = "gray"
		face.pushpull -4
		
		if($wcht>0)
			l =	entities.add_line([$length - $walk_offset - $walk_width, $width, $wcht],[$length - $walk_offset, $width, $wcht])
			l.faces[0].erase!
			l.faces[0].erase!
			l.erase!
		end
		if($concrete_height>0)
			entities.add_line([$length - $walk_offset - $walk_width, $width, $concrete_height],[$length - $walk_offset, $width, $concrete_height]).erase!

		end
	end
  end
end






#build overhead door
def create_overhead(entities)
a = $door_height
a = 0.1 if $wall_opening == 1
  trim_width = 2



if($side)
	if($side == "EW1")

		#new_face = entities.add_face([a, width - offset_length, 0],[9, width - offset_length, door_height],[9, width - offset_length - door_length, door_height],[9, width - offset_length - door_length, 0])
		
		
		
		new_face = entities.add_face([a, $width - $offset_length, 0],[a, $width - $offset_length, $door_height],[a, $width - $offset_length - $door_width, $door_height],[a, $width - $offset_length - $door_width, 0])
 		new_face.pushpull -a, true
		new_face.material = $overheadColor
    if $wall_opening 
      new_face.erase!
      entities.add_line([a, $width - $offset_length, 0],[a, $width - $offset_length - $door_width, 0]).erase!
    end
		#new_face2 = entitites.add_face([0, $width - $offset_length, 0],[0, $width - $offset_length, $door_height],[0, $width - $offset_length - $door_width, $door_height],[0, $width - $offset_length - $door_width, 0]).erase!
		new_face1 = entities.add_face([0, $width - $offset_length+trim_width, 0], [0, $width - $offset_length + trim_width, $door_height + trim_width], [0, $width - $offset_length - $door_width - trim_width, $door_height + trim_width], [0, $width - $offset_length - $door_width - trim_width, 0],[0, $width - $offset_length - $door_width, 0], [0, $width - $offset_length - $door_width, $door_height],[0, $width-$offset_length, $door_height],[0, $width-$offset_length, 0])
		new_face1.material = $faciaColor

		new_face1.pushpull 0.1, true
		
		
		


		
	
		#group_array = [new_face, new_face1, f1, f2, f3]
	#	for i in 0..($panel-1)
	#		l = entities.add_line([a, $width - $offset_length, $door_height/$panel * (i+1)],[a, $width - $door_width - $offset_length, $door_height/$panel * (i+1)])
			#group_array.push(l)
	#	end
		#group1 = entities.add_group group_array
		
		if($wcht>0)
				#entities.add_face([0, $width - $offset_length, 0],[0, $width - $offset_length, $door_height],[0, $width - $offset_length - $door_width, $door_height],[0, $width - $offset_length - $door_width, 0]).erase!
				l = entities.add_line([0,$width - $offset_length, $wcht],[0, $width - $offset_length - $door_width, $wcht])
				
				l.faces[0].erase!
				l.faces[0].erase!
				l.erase!
		end
		if($concrete_height>0)
				entities.add_line([0,$width - $offset_length,$concrete_height],[0, $width - $offset_length - $door_width, $concrete_height]).erase!
		end
		if($overhead_window>0)
			space = ($door_width-$overhead_window*$overhead_window_width)/($overhead_window+1)
			h = ($door_height*2/(($door_height/24).to_i))+4
			
			o = Geom::Point3d.new(a,$width-$offset_length-space,h)
			v1 = Geom::Vector3d.new(0,0,$overhead_window_height)
			v2 = Geom::Vector3d.new(0,-$overhead_window_width,0)
			v3 = Geom::Vector3d.new(0,-space,0)
			for i in 1..$overhead_window
				glass = entities.add_face(o,o+v1,o+v1+v2,o+v2)
				glass.material = "[Translucent_Glass_Blue]"
				glass.back_material = "[Translucent_Glass_Blue]"
				o = o+v2+v3
			end
		end
		
		
		
		
				
		f1 = entities.add_face([0, $width-$offset_length, 0], [a, $width-$offset_length, 0], [a, $width - $offset_length, $door_height], [0, $width- $offset_length, $door_height])
		f1.back_material = $faciaColor
				f1.pushpull -0.1
		f2 = entities.add_face([0, $width- $offset_length, $door_height], [a, $width- $offset_length, $door_height], [a, $width- $offset_length - $door_width, $door_height],[0, $width- $offset_length - $door_width, $door_height])
		f2.back_material = $faciaColor
		f2.pushpull -0.1
		f3 = entities.add_face([a, $width- $offset_length - $door_width, $door_height],[0, $width- $offset_length - $door_width, $door_height],[0, $width- $offset_length - $door_width, 0],[a, $width- $offset_length - $door_width, 0])
		f3.back_material = $faciaColor
		
		f3.pushpull -0.1
		entities.add_line([0,$width-$offset_length-0.1,0],[0,$width-$offset_length-$door_width+0.1,0]).erase!
		entities.add_line([2+$post_length,$width-$offset_length-0.1,0],[2+$post_length,$width-$offset_length-$door_width+0.1,0]).erase!
    if($wall_opening == 1)
      entities.add_text(($door_width/12).to_i.to_s+"X"+($door_height/12).to_i.to_s + " OPEN", [24,$width-$offset_length-$door_width/2,-2])
    elsif($overhead_opening==1)
      entities.add_text(($door_width/12).to_i.to_s+"X"+($door_height/12).to_i.to_s + " OO", [24,$width-$offset_length-$door_width/2,-2])
    else  
		  entities.add_text(($door_width/12).to_i.to_s+"X"+($door_height/12).to_i.to_s + " OHD", [24,$width-$offset_length-$door_width/2,-2])
		end
	end

	if($side == "EW2")
		#new_face = entities.add_face([$length-9, , 0],[$length-9, width - offset_length, door_height],[$length-9, width - offset_length - door_length, door_height],[9, width - offset_length - door_length, 0])
		new_face = entities.add_face([$length-a, $door_width + $offset_length, 0],[$length-a, $door_width + $offset_length, $door_height],[$length-a, $offset_length, $door_height],[$length-a, $offset_length, 0])
 		new_face.pushpull +a, true
		new_face.material = $overheadColor
    if $wall_opening == 1
      new_face.erase!
      entities.add_line([$length-a, $door_width + $offset_length, 0],[$length-a, $offset_length, 0]).erase! 
    end
		new_face = entities.add_face([$length, $offset_length-trim_width, 0], [$length, $offset_length - trim_width, $door_height + trim_width], [$length, $offset_length + $door_width + trim_width, $door_height + trim_width], [$length, $offset_length + $door_width + trim_width, 0],[$length, $offset_length + $door_width, 0], [$length, $offset_length + $door_width, $door_height],[$length, $offset_length, $door_height],[$length, $offset_length, 0])
		new_face.material = $faciaColor
		
		
		#for i in 0..($panel-1)
		#	entities.add_line([$length-a, $door_width + $offset_length, $door_height/$panel * (i+1)],[$length-a, $offset_length, $door_height/$panel * (i+1)])
		#end


		if($wcht>0)
			l =	entities.add_line([$length,$offset_length, $wcht],[$length, $offset_length + $door_width, $wcht])
				
				
			l.faces[0].erase!
			l.faces[0].erase!
			l.erase!
		end
		if($concrete_height>0)
			entities.add_line([$length,$offset_length, $concrete_height],[$length, $offset_length + $door_width, $concrete_height]).erase!
		end
		if($overhead_window>0)
			space = ($door_width-$overhead_window*$overhead_window_width)/($overhead_window+1)
			h = ($door_height*2/(($door_height/24).to_i))+4
			
			o = Geom::Point3d.new($length-a,$offset_length+space,h)
			v1 = Geom::Vector3d.new(0,0,$overhead_window_height)
			v2 = Geom::Vector3d.new(0,+$overhead_window_width,0)
			v3 = Geom::Vector3d.new(0,+space,0)
			for i in 1..$overhead_window
				glass = entities.add_face(o,o+v1,o+v1+v2,o+v2)
				glass.material = "[Translucent_Glass_Blue]"
				glass.back_material = "[Translucent_Glass_Blue]"
				o = o+v2+v3
			end
		end
		
		new_face.pushpull 0.1, true		
		left = entities.add_face([$length, $offset_length, 0], [$length-a, $offset_length, 0], [$length-a, $offset_length, $door_height], [$length, $offset_length, $door_height])
		left.back_material = $faciaColor
		entities.add_face([$length, $offset_length, $door_height], [$length-a, $offset_length, $door_height], [$length-a, $offset_length + $door_width, $door_height],[$length, $offset_length + $door_width, $door_height]).reverse!.material = $faciaColor
		right = entities.add_face([$length-a, $offset_length + $door_width, $door_height],[$length, $offset_length + $door_width, $door_height],[$length, $offset_length + $door_width, 0],[$length-a, $offset_length + $door_width, 0])
		right.back_material = $faciaColor
		left.pushpull -0.1
		right.pushpull -0.1
		
		entities.add_line([$length,$offset_length+0.1,0],[$length,$offset_length+$door_width-0.1,0]).erase!
		entities.add_line([$length-2-$post_length,$offset_length+0.1,0],[$length-2-$post_length,$offset_length+$door_width-0.1,0]).erase!

    if($wall_opening == 1)
        c = " OPEN"
    elsif($overhead_opening==1)
        c = " OO"
    else  
        c = " OHD"
    end
		entities.add_text(($door_width/12).to_i.to_s+"X"+($door_height/12).to_i.to_s + c, [$length-$door_height+24,$offset_length+$door_width/2,-2])

	end

	if($side == "SW1")
		new_face = entities.add_face([$offset_length, a, 0],[$offset_length, a, $door_height],[$offset_length + $door_width, a, $door_height],[$offset_length + $door_width, a, 0])
 		new_face.pushpull -a, true
		new_face.material = $overheadColor
    if $wall_opening
      new_face.erase!
      entities.add_line([$offset_length, a, 0],[$offset_length + $door_width, a, 0]).erase!
    end


		new_face = entities.add_face([$offset_length - trim_width, 0, 0], [$offset_length - trim_width, 0, $door_height + trim_width], [$offset_length + $door_width + trim_width, 0, $door_height + trim_width], [$offset_length + $door_width + trim_width, 0, 0], [$offset_length + $door_width, 0, 0], [$offset_length + $door_width, 0, $door_height],[$offset_length, 0, $door_height], [$offset_length, 0, 0])
		new_face.material = $faciaColor
		
		left = entities.add_face([$offset_length, 0, 0],[$offset_length, 0, $door_height],[$offset_length, a, $door_height],[$offset_length, a, 0])
		left.back_material = $faciaColor
		
		entities.add_face([$offset_length, 0, $door_height],[$offset_length + $door_width, 0, $door_height],[$offset_length + $door_width, a, $door_height],[$offset_length, a, $door_height]).back_material = $faciaColor
		right = entities.add_face([$offset_length + $door_width, 0, $door_height],[$offset_length + $door_width, 0, 0],[$offset_length + $door_width, a, 0], [$offset_length + $door_width, a, $door_height])
		right.back_material = $faciaColor

		for i in 0..($panel-1)
			entities.add_line([$offset_length,a,$door_height/$panel * (i+1)],[$offset_length+$door_width,a, $door_height/$panel * (i+1)])
		end
		
		if($wcht>0)
			l =	entities.add_line([$offset_length, 0, $wcht],[$offset_length + $door_width, 0, $wcht])
			l.faces[0].erase!
			l.faces[0].erase!
			l.erase!
		end
		if($concrete_height>0)
			entities.add_line([$offset_length, 0, $concrete_height],[$offset_length + $door_width, 0, $concrete_height]).erase!
		end
		if($overhead_window>0)
			space = ($door_width-$overhead_window*$overhead_window_width)/($overhead_window+1)
			h = ($door_height*2/(($door_height/24).to_i))+4
			
			o = Geom::Point3d.new($offset_length+space,a,h)
			v1 = Geom::Vector3d.new(0,0,$overhead_window_height)
			v2 = Geom::Vector3d.new($overhead_window_width,0,0)
			v3 = Geom::Vector3d.new(space,0,0)
			for i in 1..$overhead_window
				glass = entities.add_face(o,o+v1,o+v1+v2,o+v2)
				glass.material = "[Translucent_Glass_Blue]"
				glass.back_material = "[Translucent_Glass_Blue]"
				o = o+v2+v3
			end
		end
		new_face.pushpull 0.1
		left.pushpull -0.1
		right.pushpull -0.1
		
		entities.add_line([$offset_length+0.1,0,0],[$offset_length+$door_width-0.1,0,0]).erase!
		entities.add_line([$offset_length+0.1,2+$post_length,0],[$offset_length+$door_width-0.1,2+$post_length,0]).erase!
        if($wall_opening == 1)
        c = " OPEN"
    elsif($overhead_opening==1)
        c = " OO"
    else  
        c = " OHD"
    end
		entities.add_text(($door_width/12).to_i.to_s+"X"+($door_height/12).to_i.to_s + c, [$offset_length+24,$door_height/2,-2])
		
	end


	if($side == "SW2")
		new_face = entities.add_face([$length - $offset_length, $width - a, 0],[$length - $offset_length, $width - a, $door_height],[$length - $offset_length - $door_width, $width - a, $door_height],[$length - $offset_length - $door_width, $width - a, 0])
 		new_face.pushpull -a, true
		new_face.material = $overheadColor
    if $wall_opening
      new_face.erase!
      entities.add_line([$length - $offset_length, $width - a, 0],[$length - $offset_length - $door_width, $width - a, 0])
    end
		new_face = entities.add_face([$length - $offset_length + trim_width, $width, 0], [$length - $offset_length + trim_width, $width, $door_height + trim_width], [$length - $offset_length - trim_width - $door_width, $width, $door_height + trim_width], [$length-($offset_length + $door_width + trim_width), $width, 0], [$length-($offset_length + $door_width), $width, 0], [$length-($offset_length + $door_width), $width, $door_height],[$length - $offset_length, $width, $door_height], [$length - $offset_length, $width, 0])
		new_face.material = $faciaColor
		
		left = entities.add_face([$length - $offset_length, $width, 0],[$length - $offset_length, $width, $door_height],[$length-$offset_length, $width-a, $door_height],[$length - $offset_length, $width-a, 0])
		left.back_material = $faciaColor
		entities.add_face([$length-$offset_length, $width, $door_height],[$length-($offset_length + $door_width), $width, $door_height],[$length-($offset_length + $door_width), $width-a, $door_height],[$length-$offset_length, $width-a, $door_height]).reverse!.material = $faciaColor
		right = entities.add_face([$length-($offset_length + $door_width), $width, $door_height],[$length-($offset_length + $door_width), $width, 0],[$length-($offset_length + $door_width), $width-a, 0], [$length-($offset_length + $door_width), $width-a, $door_height])
		right.back_material = $faciaColor

		for i in 0..($panel-1)
			entities.add_line([$length-($offset_length + $door_width), $width - a,$door_height/$panel * (i+1)],[$length-$offset_length, $width - a, $door_height/$panel * (i+1)])
		end
		
		if($wcht>0)
			l =	entities.add_line([$length - $offset_length - $door_width, $width, $wcht],[$length - $offset_length, $width, $wcht])
			l.faces[0].erase!
			l.faces[0].erase!
			l.erase!
		end
		if($concrete_height>0)
			entities.add_line([$length - $offset_length - $door_width, $width, $concrete_height],[$length - $offset_length, $width, $concrete_height]).erase!
		end
		if($overhead_window>0)
			space = ($door_width-$overhead_window*$overhead_window_width)/($overhead_window+1)
			h = ($door_height*2/(($door_height/24).to_i))+4
			
			o = Geom::Point3d.new($length-$offset_length-space,$width-a,h)
			v1 = Geom::Vector3d.new(0,0,$overhead_window_height)
			v2 = Geom::Vector3d.new(-$overhead_window_width,0,0)
			v3 = Geom::Vector3d.new(-space,0,0)
			for i in 1..$overhead_window
				glass = entities.add_face(o,o+v1,o+v1+v2,o+v2)
				glass.material = "[Translucent_Glass_Blue]"
				glass.back_material = "[Translucent_Glass_Blue]"
				o = o+v2+v3
			end
		end
		
		new_face.pushpull 0.1
		left.pushpull -0.1
		right.pushpull -0.1
		
		entities.add_line([$length-$offset_length-0.1,$width,0],[$length-$offset_length-$door_width+0.1,$width,0]).erase!
		entities.add_line([$length-$offset_length-0.1,$width-2-$post_length,0],[$length-$offset_length-$door_width+0.1,$width-2-$post_length,0]).erase!
        if($wall_opening == 1)
        c = " OPEN"
    elsif($overhead_opening==1)
        c = " OO"
    else  
        c = " OHD"
    end
		entities.add_text(($door_width/12).to_i.to_s+"X"+($door_height/12).to_i.to_s + c, [$length-$offset_length-$door_width+24,$width-$door_height/2,-2])
	end
  end
end


#build interior wall
def create_interior(entities)
if($interior_side == "right")
	i = entities.add_group
	f = i.entities.add_face [$length,1,0],[$length, 1,$height],[$length, $width-1, $height],[$length,$width-1,0]
	f.back_material = $interior_color
	f.material = $interior_color
	irene = $wcht
	asmaa = $concrete_height
	$interior_walk.each do |walkdoor|
		$wcht=0
		$concrete_height = 0
	$walkside = "EW2"
	$walk_height = walkdoor[0]
	$walk_width = walkdoor[1]
	$walk_color = walkdoor[2]
	$walk_offset = walkdoor[3]
	$walk_window = walkdoor[4]
	$walk_grid = walkdoor[5]
	$door_swing = walkdoor[6]
	create_walkdoor(i.entities)
	end 
	$interior_over.each do |overhead|
	$wcht = 0
	$concrete_height = 0
	$offset_length = overhead[0]
	$door_height = overhead[1]
	$door_width = overhead[2]
	$panel = overhead[3]
	$side = "EW2"
	$overheadColor = overhead[4]
	create_overhead(i.entities)
	end
	$concrete_height = asmaa
	$wcht = irene
	t = Geom::Transformation.new [-$length+$interior_distance,0,0]
	entities.transform_entities t, i
end

if($interior_side == "left")
	i = entities.add_group
	f = i.entities.add_face [0,1,0],[0, 1,$height],[0, $width-1, $height],[0,$width-1,0]
	f.back_material = $interior_color
	f.material = $interior_color
	irene = $wcht
	asmaa = $concrete_height
		$interior_walk.each do |walkdoor|
		$wcht=0
		$concrete_height = 0
	$walkside = "EW1"
	$walk_height = walkdoor[0]
	$walk_width = walkdoor[1]
	$walk_color = walkdoor[2]
	$walk_offset = walkdoor[3]
	$walk_window = walkdoor[4]
	$walk_grid = walkdoor[5]
	$door_swing = walkdoor[6]
	create_walkdoor(i.entities)
	end 
	$interior_over.each do |overhead|
	$wcht = 0
	$concrete_height = 0
	$offset_length = overhead[0]
	$door_height = overhead[1]
	$door_width = overhead[2]
	$panel = overhead[3]
	$side = "EW1"
	$overheadColor = overhead[4]
	create_overhead(i.entities)
	end
	$wcht = irene
	$concrete_height = asmaa
	t = Geom::Transformation.new [$interior_distance,0,0]
	entities.transform_entities t, i
end
end





def draw_rectangle(side,offset,height,length,entities,h)
	if(side == "EW1")
		o = Geom::Point3d.new(0, $width-offset-length,h)
		v = Geom::Vector3d.new(0,length, 0)
		h = Geom::Vector3d.new(0,0,height)
		p = 0.5
	end
	if(side == "SW2")
		o = Geom::Point3d.new($length-length-offset,$width,h)
		v = Geom::Vector3d.new(length,0,0)
		h = Geom::Vector3d.new(0,0,height)
		p = 0.5
	end
	
	if(side == "SW1")
		o = Geom::Point3d.new(offset,0,h)
		v = Geom::Vector3d.new(length,0,0)
		h = Geom::Vector3d.new(0,0,height)
		p = 0.5
	end
	
	if(side == "EW2")
		o = Geom::Point3d.new($length,offset,h)
		v = Geom::Vector3d.new(0,length,0)
		h = Geom::Vector3d.new(0,0,height)
		p = 0.5
	end
	
	p1 = o
	p2 = o + v
	p3 = o + v + h
	p4 = o + h
	entities.add_face(p1,p2,p3,p4).pushpull -p
end



def build_concrete(entities)
	concrete_group = entities.add_group
	concrete = concrete_group.entities.add_face([0,0,3],[0,-$concrete_length,3],[$concrete_width,-$concrete_length,3],[$concrete_width,0,3])
	concrete.material = "white"

	if($concrete_side == "SW1")
		# r = Geom::Transformation.rotation [$porch_width/2, 0, 0], [0,0,1], 180.degrees
		# t = Geom::Transformation.new [$porch_offset,0, $porch_height]
		# entities.transform_entities r, porch_group
		# entities.transform_entities t, porch_group
		t = Geom::Transformation.new [$concrete_offset, 0, -2.9]
		entities.transform_entities t, concrete_group
	end
	
	if($concrete_side == "SW2")
		r = Geom::Transformation.rotation [0,0,0],[0,0,1], 180.degrees
		t = Geom::Transformation.new [$length-$concrete_offset, $width, -2.9]
		entities.transform_entities r, concrete_group
		entities.transform_entities t, concrete_group
	end
	
	if($concrete_side == "EW1")
		r = Geom::Transformation.rotation [0,0,0],[0,0,1], 270.degrees
		t = Geom::Transformation.new [0, $width-$concrete_offset, -2.9]
		entities.transform_entities r, concrete_group
		entities.transform_entities t, concrete_group
	end
	
	if($concrete_side == "EW2")
		r = Geom::Transformation.rotation [0,0,0],[0,0,1], 90.degrees
		t = Geom::Transformation.new [$length, $concrete_offset, -2.9]
		entities.transform_entities r, concrete_group
		entities.transform_entities t, concrete_group
	end
	
		concrete.pushpull -2
end





# build hydraulic
def build_hydraulic(entities, r, t)
hydraulic_group = entities.add_group
hydraulic_group.entities.add_text(($hydraulic_width/12).to_i.to_s+"X"+($hydraulic_height/12).to_i.to_s + " HYDRO", [-$hydraulic_width/2,-20,-2])
f = hydraulic_group.entities.add_face [0,0,6],[0,-4, 6],[0,-4,6+$hydraulic_height],[0,0,8+6+$hydraulic_height]
b = hydraulic_group.entities.add_line([0,0,6+$hydraulic_height],[0,-4,6+$hydraulic_height])
b.faces[0].material = "black"

if ($hydraulic_wainscot >0 and $hydraulic_wainscot < $hydraulic_height)
	w = hydraulic_group.entities.add_line([0,0,6+$hydraulic_wainscot],[0,-4,6+$hydraulic_wainscot]) 
	b.faces[0].pushpull $hydraulic_width, true
	b.faces[0].back_material = "black"
	w.find_faces
	w.faces[0].material = $hydraulic_color

	w.faces[1].material = $hydraulic_wainscot_color
	w.faces[1].pushpull $hydraulic_width, true
	w.faces[1].back_material = $hydraulic_wainscot_color
	w.faces[0].pushpull $hydraulic_width, true
	w.faces[0].back_material = $hydraulic_color
end
if ($hydraulic_wainscot == $hydraulic_height)
	b.faces[1].material = $hydraulic_wainscot_color
	b.faces[1].pushpull $hydraulic_width, true
	b.faces[1].back_material = $hydraulic_wainscot_color
	b.faces[0].pushpull $hydraulic_width, true
	b.faces[0].back_material = "black"
end
if ($hydraulic_wainscot == 0)
	b.faces[1].material = $hydraulic_color
	b.faces[1].pushpull $hydraulic_width, true
	b.faces[1].back_material = $hydraulic_wainscot_color
	b.faces[0].pushpull $hydraulic_width, true
	b.faces[0].back_material = "black"
end



entities.transform_entities r, hydraulic_group
entities.transform_entities t, hydraulic_group

end




#build hydraulic and move
def create_hydraulic(entities)
 # r = Geom::Transformation.rotation [$porch_width/2, 0, 0], [0,0,1], 180.degrees
		# t = Geom::Transformation.new [$porch_offset,0, $porch_height]
		# entities.transform_entities r, porch_group
		# entities.transform_entities t, porch_group
	if($hydraulic_side == "SW1")
		r = Geom::Transformation.rotation [0, 0, 0], [0,0,1], 360.degrees
		t = Geom::Transformation.new [$hydraulic_offset+$hydraulic_width,0, -6]
		build_hydraulic(entities,r,t)
	end
	if($hydraulic_side == "SW2")
		r = Geom::Transformation.rotation [0, 0, 0], [0,0,1], 180.degrees
		t = Geom::Transformation.new [$length-$hydraulic_offset-$hydraulic_width,$width, -6]
		build_hydraulic(entities,r,t)
	end
	if($hydraulic_side == "EW2")
		r = Geom::Transformation.rotation [0, 0, 0], [0,0,1], 90.degrees
		t = Geom::Transformation.new [$length,$hydraulic_offset+$hydraulic_width, -6]
		build_hydraulic(entities,r,t)
	end
	if($hydraulic_side == "EW1")
		r = Geom::Transformation.rotation [0, 0, 0], [0,0,1], 270.degrees
		t = Geom::Transformation.new [0,$width-$hydraulic_offset-$hydraulic_width, -6]
		build_hydraulic(entities,r,t)
	end
 end
$size1 = 10
if($size1>0)
	for i in 1..$size1
		$hydraulic_wainscot = worksheet.Cells(79,1+i).Value.to_f
		$hydraulic_height = worksheet.Cells(80,1+i).Value.to_f
		$hydraulic_width = worksheet.Cells(81,1+i).Value.to_f
		$hydraulic_wainscot_color = worksheet.Cells(82,1+i).Value
		$hydraulic_color = worksheet.Cells(83,1+i).Value
		$hydraulic_side = worksheet.Cells(84,1+i).Value
		$hydraulic_offset = worksheet.Cells(85,1+i).Value.to_f
		if($hydraulic_offset>0)
			create_hydraulic(entities)
		end
	end
end

 # build GB logo
def add_3d_letter_ew1(entities, string, yorigin, zorigin)
  logo_group = entities.add_group

  xaxis = Geom::Vector3d.new(0, 1, 0)
  yaxis = Geom::Vector3d.new(0, 0, 1)
  zaxis = Geom::Vector3d.new(1, 0, 0)

  logo_origin = Geom::Point3d.new 0, yorigin, zorigin
  transform_logo_group = Geom::Transformation.axes logo_origin, xaxis, yaxis, zaxis

  logo_group.transform! transform_logo_group
  logo_group.transform! Geom::Transformation.rotation(logo_origin, yaxis, Math::PI)

  # Add G first
  logo_group.entities.add_3d_text(string, TextAlignCenter, "Arial", false, false, 10.0, 0.0, 0, true, 1.0) 
  logo_group.material = $GBcolor
end

#build post
def build_post(entities)
	x = 1.5
	
	if($post_side == "SW1")
		if($post_x ==1)
			entities.add_line [$post_offset,1,-1],[$post_offset + $post_length,1+$post_width,-1]
			entities.add_line [$post_offset + $post_length,1,-1],[$post_offset,1+$post_width,-1]
		else
			a = entities.add_face [$post_offset,1,-1],[$post_offset + $post_length,1,-1],[$post_offset + $post_length,1+$post_width,-1],[$post_offset,1+$post_width,-1]
			a.reverse!
		end
	end
	
	if($post_side == "SW2")
		if($post_x == 1)
			entities.add_line [$length-$post_offset,$width-1,-1],[$length-$post_offset - $post_length,$width-1-$post_width,-1]
			entities.add_line [$length-$post_offset - $post_length,$width-1,-1],[$length-$post_offset,$width-1-$post_width,-1]
		else
			a = entities.add_face [$length-$post_offset,$width-1,-1],[$length-$post_offset - $post_length,$width-1,-1],[$length-$post_offset - $post_length,$width-1-$post_width,-1],[$length-$post_offset,$width-1-$post_width,-1]
			a.reverse!
		end
		
	end
	
	if($post_side == "EW1")
		if($post_x == 1)
			entities.add_line [1,$width-$post_offset,-1],[1+$post_width,$width-$post_offset-$post_length,-1]
			entities.add_line [1+$post_width,$width-$post_offset,-1],[1,$width-$post_offset-$post_length,-1]
		else
			entities.add_face [1,$width-$post_offset,-1],[1+$post_width,$width-$post_offset,-1],[1+$post_width,$width-$post_offset-$post_length,-1],[1,$width-$post_offset-$post_length,-1]
		end
		a = $width/12/6
		pt = Geom::Point3d.new($left-($length*2*x/8/12), $width - $post_offset + $width*x/4/12, -1)
		v1 = Geom::Vector3d.new(a*0.8*x,0,0)
		v2 = Geom::Vector3d.new(0,a*x,0)
		entities.add_circle pt, [0,0,1], a*x
		entities.add_edges(pt-v2-v1-v1,pt-v2+v1+v1,pt-v2-v2+v1+v1,pt-v2-v2-v1-v1,pt-v2-v1-v1)
		entities.add_text $square.to_s, pt + Geom::Vector3d.new(-$length*x/18/12,-a*x-($width*x/12/12),0) 
		entities.add_text $circle.to_s, pt + Geom::Vector3d.new(-$length*x/25/12,1,0)
	end
	
	if($post_side == "EW2")
		if($post_x == 1)
			entities.add_line [$length-1,$post_offset,-1],[$length-1-$post_width,$post_offset+$post_length,-1]
			entities.add_line [$length-1-$post_width,$post_offset,-1],[$length-1,$post_offset+$post_length,-1]
		else
			entities.add_face [$length-1,$post_offset,-1],[$length-1-$post_width,$post_offset,-1],[$length-1-$post_width,$post_offset+$post_length,-1],[$length-1,$post_offset+$post_length,-1]
		end
		
		d = $right
		a = $width/12/6
		v1 = Geom::Vector3d.new(a*x*0.8,0,0)
		v2 = Geom::Vector3d.new(0,a*x,0)
		pt = Geom::Point3d.new($length*1.05+d, $post_offset+$width*x/4/12, -1)
		entities.add_circle pt, [0,0,1],a*x
		entities.add_edges(pt-v2-v1-v1,pt-v2+v1+v1,pt-v2-v2+v1+v1,pt-v2-v2-v1-v1,pt-v2-v1-v1)
		entities.add_text $square.to_s, pt + Geom::Vector3d.new(-$length*x/18/12,-a*x-($width*x/12/12),0) 
		entities.add_text $circle.to_s, pt + Geom::Vector3d.new(-$length*x/25/12,1,0)
	end
end

#build GB logo
def add_3d_letter_ew2(entities, string, yorigin, zorigin, xorigin)
  logo_group = entities.add_group

  xaxis = Geom::Vector3d.new(0, 1, 0)
  yaxis = Geom::Vector3d.new(0, 0, 1)
  zaxis = Geom::Vector3d.new(1, 0, 0)

  logo_origin = Geom::Point3d.new xorigin, yorigin, zorigin 
  transform_logo_group = Geom::Transformation.axes logo_origin, xaxis, yaxis, zaxis

  logo_group.transform! transform_logo_group

  # Add G first
  logo_group.entities.add_3d_text(string, TextAlignCenter, "Arial", false, false, 10.0, 0.0, 0, true, 1.0)
	logo_group.material = $GBcolor
end


# build cupola
def build_cupola(a, entities, pitch, t, roof, wall)
		
		cupola_group = entities.add_group
		c = cupola_group.entities.add_face([0,0,0],[a,0,0],[a,a,0],[0,a,0])
		c.erase!
		cupola_group.entities.add_text (a/12).to_i.to_s+"x"+(a/12).to_i.to_s+ " Cupola", [0,-10,0]
		
=begin
		f1 = cupola_group.entities.add_face([a,0,0],[a,a/2, a*$pitch/2/12], [a, a, 0], [a, a, a*$pitch/2/12 + 4 + a], [a,0,a*$pitch/2/12 + 4 + a])
		f1.pushpull -a

		cupola_group.entities.add_edges([0,0,a*$pitch/2/12 + 4],[a,0,a*$pitch /2/12+ 4 ])[0].faces[1].material = roof
		cupola_group.entities.add_edges([0,0,a*$pitch/2/12 + 4],[a,0,a*$pitch/2/12 + 4 ])[0].faces[0].material = wall
		cupola_group.entities.add_edges([a,0,a*$pitch/2/12 + 4],[a,a,a*$pitch/2/12 + 4 ])[0].faces[1].material = roof
		cupola_group.entities.add_edges([a,0,a*$pitch/2/12 + 4],[a,a,a*$pitch/2/12 + 4 ])[0].faces[0].material = wall
		cupola_group.entities.add_edges([a,a,a*$pitch/2/12 + 4],[0,a,a*$pitch/2/12 + 4])[0].faces[1].material = roof
		cupola_group.entities.add_edges([a,a,a*$pitch/2/12 + 4],[0,a,a*$pitch/2/12 + 4])[0].faces[0].material = wall
		cupola_group.entities.add_edges([0,0,a*$pitch/2/12 + 4],[0,a,a*$pitch/2/12 + 4])[0].faces[0].material = roof
		cupola_group.entities.add_edges([0,0,a*$pitch/2/12 + 4],[0,a,a*$pitch/2/12 + 4])[0].faces[1].material = wall
		
		face = cupola_group.entities.add_face([0,0,a*$pitch/2/12 + 4 + a],[a/2,a/2, a*3/2+ a*$pitch/2/12 + 4], [0,a,a*$pitch/2/12 + 4 + a])
		face.material = roof
		face = cupola_group.entities.add_face([0,a,a*$pitch/2/12 + 4 + a],[a/2,a/2, a*3/2 + a*$pitch/2/12 + 4],[a,a,a*$pitch/2/12 + 4 + a])
		face.material = roof
		face = cupola_group.entities.add_face([a,a,a*$pitch/2/12 + 4 + a],[a/2,a/2, a*3/2 + a*$pitch/2/12 + 4],[a,0,a*$pitch/2/12 + 4 + a])
		face.material = roof
		face =  cupola_group.entities.add_face([a,0,a*$pitch/2/12 + 4 + a],[a/2,a/2, a*3/2 + a*$pitch/2/12 + 4],[0,0,a*$pitch/2/12 + 4 + a])
		face.material = roof
=end
		entities.transform_entities t, cupola_group


end

  #build window
def create_window(entities)

$window_color = "white"
if $window_side == "SW1"

  if($window_height > 10*12)
    $window_height = 0
    origin = Geom::Point3d.new($window_offset,-60,$window_height) 
     v1 = Geom::Vector3d.new(0,3,0)
     v2 = Geom::Vector3d.new($window_width,0,0)
     entities.add_line(origin+v1,origin+v1+v2)
     entities.add_line(origin+v1+v1,origin)
     entities.add_line(origin+v2,origin+v2+v1+v1)
     entities.add_line(origin+v1+v1,origin+v1+v1+v2)
     entities.add_line(origin,origin+v2)
  
     entities.add_text ($window_length/12).to_i.to_s+" X "+($window_width/12).to_i.to_s+ " WIN", origin-v1-v1

  else
    $window_height = 0
	   origin = Geom::Point3d.new($window_offset,0,$window_height) 
	   v1 = Geom::Vector3d.new(0,3,0)
	   v2 = Geom::Vector3d.new($window_width,0,0)
	   entities.add_line(origin+v1,origin+v1+v2)
	   entities.add_line(origin+v1+v1,origin)
	   entities.add_line(origin+v2,origin+v2+v1+v1)
	   entities.add_line(origin+v1+v1,origin+v1+v1+v2)
	
	   entities.add_text ($window_length/12).to_i.to_s+" X "+($window_width/12).to_i.to_s+ " WIN", origin-v1-v1
	end
=begin
	origin = Geom::Point3d.new($window_offset,0,$window_height)
	v1 = Geom::Vector3d.new(0,0,$window_length)
	v2 = Geom::Vector3d.new($window_width,0,0)
	p1 = origin + v1
	p2 = origin + v1 + v2
	p3 = origin + v2
	a = entities.add_face origin, p1, p2, p3
	a.back_material = $window_color
	a.material = $window_color

	origin = origin + Geom::Vector3d.new(2,0,2)
	v1 = Geom::Vector3d.new(0,0, $window_length -4)
	v2 = Geom::Vector3d.new($window_width -4,0,0)  
	p1 = origin + v1
	p2 = origin + v1 + v2
	p3 = origin + v2
	b = entities.add_face origin, p1, p2, p3
	b.material = "[Translucent_Glass_Blue]"
	b.back_material = "[Translucent_Glass_Blue]"
	
	window_array = [a,b]
	
	
	if($window_type == "Verticle")
		p = origin
		
		origin = origin + Geom::Vector3d.new(0.1,-0.1,0)
		v1 = Geom::Vector3d.new(0,0,($window_length-4)/2)
		v2 = Geom::Vector3d.new($window_width-4,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		c = entities.add_face origin, p1, p2, p3
		c.back_material = $window_color
		window_array.push(c)

		origin = origin + Geom::Vector3d.new(2,0,2)
		v1 = Geom::Vector3d.new(0,0,(($window_length-4)/2) - 4)
		v2 = Geom::Vector3d.new($window_width-4-4,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		d = entities.add_face origin, p1, p2, p3
		d.back_material = "[Translucent_Glass_Blue]"
		d.material = "[Translucent_Glass_Blue]"
		window_array.push(d)
		
		
		
		origin = p + Geom::Vector3d.new(0,-1.4,($window_length-4)/2)
		v1 = Geom::Vector3d.new(0,0,($window_length-4)/2)
		v2 = Geom::Vector3d.new($window_width-4,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		e = entities.add_face origin, p1, p2, p3
		e.back_material = $window_color
		window_array.push(e)

		origin = origin + Geom::Vector3d.new(2,0,2)
		v1 = Geom::Vector3d.new(0,0,(($window_length-4)/2) - 4)
		v2 = Geom::Vector3d.new($window_width-4-4,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		f = entities.add_face origin, p1, p2, p3
		f.back_material = "[Translucent_Glass_Blue]"
		f.material = "[Translucent_Glass_Blue]"
		window_array.push(f)
	end
	if($window_type == "Slider")
		p = origin
		
		origin = origin + Geom::Vector3d.new(0,-0.1,0)
		v1 = Geom::Vector3d.new(0,0,$window_length-4)
		v2 = Geom::Vector3d.new(($window_width-4)/2,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		c = entities.add_face origin, p1, p2, p3
		c.back_material = $window_color
		window_array.push(c)

		origin = origin + Geom::Vector3d.new(2,0,2)
		v1 = Geom::Vector3d.new(0,0,$window_length-4 - 4)
		v2 = Geom::Vector3d.new(($window_width-4)/2-4,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		d = entities.add_face origin, p1, p2, p3
		d.back_material = "[Translucent_Glass_Blue]"
		d.material = "[Translucent_Glass_Blue]"
		window_array.push(d)
		
		
		
		origin = p + Geom::Vector3d.new(($window_width-4)/2,-1.4,0)
		v1 = Geom::Vector3d.new(0,0,$window_length-4)
		v2 = Geom::Vector3d.new(($window_width-4)/2,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		e = entities.add_face origin, p1, p2, p3
		e.back_material = $window_color
		window_array.push(e)
		
		
		
		origin = origin + Geom::Vector3d.new(2,0,2)
		v1 = Geom::Vector3d.new(0,0,$window_length-4 - 4)
		v2 = Geom::Vector3d.new(($window_width-4)/2-4,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		f = entities.add_face origin, p1, p2, p3
		f.back_material = "[Translucent_Glass_Blue]"
		f.material = "[Translucent_Glass_Blue]"
		window_array.push(f)
	end
	
	if($window_grid == 1)
		origin = Geom::Point3d.new($window_offset,0,$window_height) + Geom::Vector3d.new(2,0.1,2)
		window_array.push(entities.add_line(origin + Geom::Vector3d.new(($window_width -4)/4,0,0), origin + Geom::Vector3d.new(($window_width -4)/4,0,0) + Geom::Vector3d.new(0,0, $window_length -4)))
		window_array.push(entities.add_line(origin + Geom::Vector3d.new(($window_width -4)*3/4,0,0), origin + Geom::Vector3d.new(($window_width -4)*3/4,0,0) + Geom::Vector3d.new(0,0, $window_length -4)))
		window_array.push(entities.add_line(origin + Geom::Vector3d.new(0,0,($window_length-4)/4), origin + Geom::Vector3d.new(0,0,($window_length-4)/4) +  Geom::Vector3d.new($window_width -4,0,0)))
		window_array.push(entities.add_line(origin + Geom::Vector3d.new(0,0,($window_length-4)*3/4), origin + Geom::Vector3d.new(0,0,($window_length-4)*3/4) +  Geom::Vector3d.new($window_width -4,0,0)))
	end
	
	if($window_shutter == 1)
		vs = Geom::Vector3d.new(14,0,0)
		v1 = Geom::Vector3d.new(0,0,$window_length)
		v2 = Geom::Vector3d.new($window_width,0,0)
		
		o = Geom::Point3d.new($window_offset,-1,$window_height)
		shutter1 = entities.add_face o, o+v1, o+v1-vs, o-vs

		shutter2 = entities.add_face o+v2, o+v2+v1,o+v2+v1+vs,o+v2+vs

		shutter1.material = "Kynar Matte Black Trim"
		shutter2.back_material = "Kynar Matte Black Trim"
	end
	
	window_group = entities.add_group window_array
	a.pushpull 2, true
	a.pushpull 4, true
	if($window_type != "Fixed")
		c.pushpull -1.2
		e.pushpull -1.2
	end
=end


end

if $window_side == "SW2"
  if $window_height>120
    $window_height = 0
    origin = Geom::Point3d.new($length-$window_offset,$width+60,$window_height)
    v1 = Geom::Vector3d.new(0,-3,0)
    v2 = Geom::Vector3d.new(-$window_width,0,0)
    
    entities.add_line(origin+v1,origin+v1+v2)
    entities.add_line(origin+v1+v1,origin)
    entities.add_line(origin+v2,origin+v2+v1+v1)
    entities.add_line(origin+v1+v1,origin+v1+v1+v2)
    entities.add_line(origin,origin+v2)
    
    entities.add_text ($window_length/12).to_i.to_s+" X "+($window_width/12).to_i.to_s+ " WIN", origin-v1-v1+v2
  else
    $window_height = 0
  	origin = Geom::Point3d.new($length-$window_offset,$width,$window_height)
  	v1 = Geom::Vector3d.new(0,-3,0)
  	v2 = Geom::Vector3d.new(-$window_width,0,0)
  	
  	entities.add_line(origin+v1,origin+v1+v2)
  	entities.add_line(origin+v1+v1,origin)
  	entities.add_line(origin+v2,origin+v2+v1+v1)
  	entities.add_line(origin+v1+v1,origin+v1+v1+v2)
  	
  	entities.add_text ($window_length/12).to_i.to_s+" X "+($window_width/12).to_i.to_s+ " WIN", origin-v1-v1+v2
=begin
	origin = Geom::Point3d.new($length-$window_offset,$width,$window_height)
	v1 = Geom::Vector3d.new(0,0,$window_length)
	v2 = Geom::Vector3d.new(-$window_width,0,0)
	p1 = origin + v1
	p2 = origin + v1 + v2
	p3 = origin + v2
	a = entities.add_face origin, p1, p2, p3
	a.back_material = $window_color
	a.material = $window_color

	origin = origin + Geom::Vector3d.new(-2,0,2)
	v1 = Geom::Vector3d.new(0,0, $window_length -4)
	v2 = Geom::Vector3d.new(-$window_width + 4,0,0)  
	p1 = origin + v1
	p2 = origin + v1 + v2
	p3 = origin + v2
	b = entities.add_face origin, p1, p2, p3
	b.material = "[Translucent_Glass_Blue]"
	b.back_material = "[Translucent_Glass_Blue]"

	window_array = [a,b]
	
	if($window_type == "Verticle")
		p = origin
		
		origin = origin + Geom::Vector3d.new(-0.1,0.1,0)
		v1 = Geom::Vector3d.new(0,0,($window_length-4)/2)
		v2 = Geom::Vector3d.new(-$window_width+4,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		c = entities.add_face origin, p1, p2, p3
		c.back_material = $window_color
		window_array.push(c)
		
		
		origin = origin + Geom::Vector3d.new(-2,0,2)
		v1 = Geom::Vector3d.new(0,0,(($window_length-4)/2) - 4)
		v2 = Geom::Vector3d.new(-$window_width+4+4,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		d = entities.add_face origin, p1, p2, p3
		d.back_material = "[Translucent_Glass_Blue]"
		d.material = "[Translucent_Glass_Blue]"
		window_array.push(d)
		
		
		
		origin = p + Geom::Vector3d.new(0,+1.4,($window_length-4)/2)
		v1 = Geom::Vector3d.new(0,0,($window_length-4)/2)
		v2 = Geom::Vector3d.new(-$window_width+4,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		e = entities.add_face origin, p1, p2, p3
		e.back_material = $window_color
		window_array.push(e)
		
		
		origin = origin + Geom::Vector3d.new(-2,0,2)
		v1 = Geom::Vector3d.new(0,0,(($window_length-4)/2) - 4)
		v2 = Geom::Vector3d.new(-$window_width+4+4,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		f = entities.add_face origin, p1, p2, p3
		f.back_material = "[Translucent_Glass_Blue]"
		f.material = "[Translucent_Glass_Blue]"
		window_array.push(f)
	end
	if($window_type == "Slider")
		p = origin
		
		origin = origin + Geom::Vector3d.new(0,0.1,0)
		v1 = Geom::Vector3d.new(0,0,$window_length-4)
		v2 = Geom::Vector3d.new(-($window_width-4)/2,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		c = entities.add_face origin, p1, p2, p3
		c.back_material = $window_color
		window_array.push(c)
		
		
		origin = origin + Geom::Vector3d.new(-2,0,2)
		v1 = Geom::Vector3d.new(0,0,$window_length-4 - 4)
		v2 = Geom::Vector3d.new(-($window_width-4)/2+4,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		d = entities.add_face origin, p1, p2, p3
		d.back_material = "[Translucent_Glass_Blue]"
		d.material = "[Translucent_Glass_Blue]"
		window_array.push(d)
		
		
		
		origin = p + Geom::Vector3d.new(-($window_width-4)/2,+1.4,0)
		v1 = Geom::Vector3d.new(0,0,$window_length-4)
		v2 = Geom::Vector3d.new(-($window_width-4)/2,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		e = entities.add_face origin, p1, p2, p3
		e.back_material = $window_color
		window_array.push(e)

		origin = origin + Geom::Vector3d.new(-2,0,2)
		v1 = Geom::Vector3d.new(0,0,$window_length-4 - 4)
		v2 = Geom::Vector3d.new(-($window_width-4)/2+4,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		f = entities.add_face origin, p1, p2, p3
		f.back_material = "[Translucent_Glass_Blue]"
		f.material = "[Translucent_Glass_Blue]"
		window_array.push(f)
		
	end
	
	if($window_grid == 1)
		origin = Geom::Point3d.new($length-$window_offset,$width,$window_height) + Geom::Vector3d.new(-2,-0.1,2)
		window_array.push(entities.add_line(origin + Geom::Vector3d.new((-$window_width +4)/4,0,0), origin + Geom::Vector3d.new((-$window_width +4)/4,0,0) + Geom::Vector3d.new(0,0, $window_length -4)))
		window_array.push(entities.add_line(origin + Geom::Vector3d.new((-$window_width +4)*3/4,0,0), origin + Geom::Vector3d.new((-$window_width +4)*3/4,0,0) + Geom::Vector3d.new(0,0, $window_length -4)))
		window_array.push(entities.add_line(origin + Geom::Vector3d.new(0,0,($window_length-4)/4), origin + Geom::Vector3d.new(0,0,($window_length-4)/4) +  Geom::Vector3d.new(-$window_width +4,0,0)))
		window_array.push(entities.add_line(origin + Geom::Vector3d.new(0,0,($window_length-4)*3/4), origin + Geom::Vector3d.new(0,0,($window_length-4)*3/4) +  Geom::Vector3d.new(-$window_width +4,0,0)))
	end
	
	if($window_shutter == 1)
		vs = Geom::Vector3d.new(-14,0,0)
		v1 = Geom::Vector3d.new(0,0,$window_length)
		v2 = Geom::Vector3d.new(-$window_width,0,0)
		
		o = Geom::Point3d.new($length-$window_offset,$width+1,$window_height)
		shutter1 = entities.add_face o, o+v1, o+v1-vs, o-vs

		shutter2 = entities.add_face o+v2, o+v2+v1,o+v2+v1+vs,o+v2+vs

		shutter1.material = "Kynar Matte Black Trim"
		shutter2.back_material = "Kynar Matte Black Trim"
	end
	
	window_group = entities.add_group window_array
	a.pushpull 2
	if($window_type != "Fixed")
		c.pushpull -1.2
		e.pushpull -1.2
	end
=end	
  end
end

if $window_side == "EW2"
  if $window_height>120
    $window_height = 0
    origin = Geom::Point3d.new($length+60,$window_offset,$window_height)
    v1 = Geom::Vector3d.new(-3,0,0)
    v2 = Geom::Vector3d.new(0,$window_width,0)
    v3 = Geom::Vector3d.new(0,24,0)
    entities.add_line(origin+v1,origin+v1+v2)
    entities.add_line(origin+v1+v1,origin)
    entities.add_line(origin+v2,origin+v2+v1+v1)
    entities.add_line(origin+v1+v1,origin+v1+v1+v2)
    entities.add_line(origin,origin+v2)
    
    entities.add_text ($window_length/12).to_i.to_s+" X "+($window_width/12).to_i.to_s+ " WIN", origin-v1-v1+v3
  else
    $window_height = 0
    origin = Geom::Point3d.new($length,$window_offset,$window_height)
    v1 = Geom::Vector3d.new(-3,0,0)
    v2 = Geom::Vector3d.new(0,$window_width,0)
    v3 = Geom::Vector3d.new(0,24,0)
  	entities.add_line(origin+v1,origin+v1+v2)
  	entities.add_line(origin+v1+v1,origin)
  	entities.add_line(origin+v2,origin+v2+v1+v1)
  	entities.add_line(origin+v1+v1,origin+v1+v1+v2)
  	
  	entities.add_text ($window_length/12).to_i.to_s+" X "+($window_width/12).to_i.to_s+ " WIN", origin-v1-v1+v3
=begin
	origin = Geom::Point3d.new($length,$window_offset,$window_height)
	v1 = Geom::Vector3d.new(0,0,$window_length)
	v2 = Geom::Vector3d.new(0,$window_width,0)
	p1 = origin + v1
	p2 = origin + v1 + v2
	p3 = origin + v2
	a = entities.add_face origin, p1, p2, p3
	a.back_material = $window_color
	a.material = $window_color

	origin = origin + Geom::Vector3d.new(0,2,2)
	v1 = Geom::Vector3d.new(0,0, $window_length -4)
	v2 = Geom::Vector3d.new(0, $window_width - 4,0)  
	p1 = origin + v1
	p2 = origin + v1 + v2
	p3 = origin + v2
	b = entities.add_face origin, p1, p2, p3
	b.material = "[Translucent_Glass_Blue]"
	b.back_material = "[Translucent_Glass_Blue]"
	
	window_array = [a,b]
	

	
	if($window_type == "Verticle")
		p = origin
		
		origin = origin + Geom::Vector3d.new(0.1,0.1,0)
		v1 = Geom::Vector3d.new(0,0,($window_length-4)/2)
		v2 = Geom::Vector3d.new(0,+$window_width-4,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		c = entities.add_face origin, p1, p2, p3
		c.back_material = $window_color
		window_array.push(c)
		
		origin = origin + Geom::Vector3d.new(0,2,2)
		v1 = Geom::Vector3d.new(0,0,(($window_length-4)/2) - 4)
		v2 = Geom::Vector3d.new(0,$window_width-4-4,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		d = entities.add_face origin, p1, p2, p3
		d.material = "[Translucent_Glass_Blue]"
		d.back_material = "[Translucent_Glass_Blue]"
		window_array.push(d)
		
		
		origin = p + Geom::Vector3d.new(1.4,0,($window_length-4)/2)
		v1 = Geom::Vector3d.new(0,0,($window_length-4)/2)
		v2 = Geom::Vector3d.new(0,+$window_width-4,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		e = entities.add_face origin, p1, p2, p3
		e.back_material = $window_color
		window_array.push(e)
		
		
		origin = origin + Geom::Vector3d.new(0,2,2)
		v1 = Geom::Vector3d.new(0,0,(($window_length-4)/2) - 4)
		v2 = Geom::Vector3d.new(0,$window_width-4-4,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		f = entities.add_face origin, p1, p2, p3
		f.back_material = "[Translucent_Glass_Blue]"
		f.material = "[Translucent_Glass_Blue]"
		window_array.push(f)
		
		
	end
	if($window_type == "Slider")
		p = origin
		
		origin = origin + Geom::Vector3d.new(0.1,0,0)
		v1 = Geom::Vector3d.new(0,0,$window_length-4)
		v2 = Geom::Vector3d.new(0, +($window_width-4)/2,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		c = entities.add_face origin, p1, p2, p3
		c.back_material = $window_color
		
		window_array.push(c)

		origin = origin + Geom::Vector3d.new(0,2,2)
		v1 = Geom::Vector3d.new(0,0,$window_length-4 - 4)
		v2 = Geom::Vector3d.new(0, ($window_width-4)/2-4,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		d = entities.add_face origin, p1, p2, p3
		d.back_material = "[Translucent_Glass_Blue]"
		d.material = "[Translucent_Glass_Blue]"
		
		window_array.push(d)
		
		
	
		origin = p + Geom::Vector3d.new(1.4, +($window_width-4)/2,0)
		v1 = Geom::Vector3d.new(0,0,$window_length-4)
		v2 = Geom::Vector3d.new(0, +($window_width-4)/2,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		e = entities.add_face origin, p1, p2, p3
		e.back_material = $window_color
		
		window_array.push(e)

		origin = origin + Geom::Vector3d.new(0,2,2)
		v1 = Geom::Vector3d.new(0,0,$window_length-4 - 4)
		v2 = Geom::Vector3d.new(0, ($window_width-4)/2-4,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		f = entities.add_face origin, p1, p2, p3
		f.back_material = "[Translucent_Glass_Blue]"
		f.material = "[Translucent_Glass_Blue]"
		window_array.push(f)

	end
	
	if($window_grid == 1)
		origin = Geom::Point3d.new($length,$window_offset,$window_height) + Geom::Vector3d.new(-0.1,2,2)
		window_array.push(entities.add_line(origin + Geom::Vector3d.new(0,($window_width -4)/4,0), origin + Geom::Vector3d.new(0,($window_width -4)/4,0) + Geom::Vector3d.new(0,0, $window_length -4)))
		window_array.push(entities.add_line(origin + Geom::Vector3d.new(0,($window_width -4)*3/4,0), origin + Geom::Vector3d.new(0,($window_width -4)*3/4,0) + Geom::Vector3d.new(0,0, $window_length -4)))
		window_array.push(entities.add_line(origin + Geom::Vector3d.new(0,0,($window_length-4)/4), origin + Geom::Vector3d.new(0,0,($window_length-4)/4) +  Geom::Vector3d.new(0, $window_width -4,0)))
		window_array.push(entities.add_line(origin + Geom::Vector3d.new(0,0,($window_length-4)*3/4), origin + Geom::Vector3d.new(0,0,($window_length-4)*3/4) +  Geom::Vector3d.new(0, $window_width -4,0)))
	end
		
	if($window_shutter == 1)
		vs = Geom::Vector3d.new(0,14,0)
		v1 = Geom::Vector3d.new(0,0,$window_length)
		v2 = Geom::Vector3d.new(0,$window_width,0)
		
		o = Geom::Point3d.new($length+1,$window_offset,$window_height)
		shutter1 = entities.add_face o, o+v1, o+v1-vs, o-vs

		shutter2 = entities.add_face o+v2, o+v2+v1,o+v2+v1+vs,o+v2+vs

		
		shutter1.material = "Kynar Matte Black Trim"
		shutter2.back_material = "Kynar Matte Black Trim"
	end

	window_group = entities.add_group window_array
	a.pushpull 2
	if($window_type != "Fixed")
		c.pushpull -1.2
		e.pushpull -1.2
	end
=end
	end
end
if $window_side == "EW1"
  if $window_height > 120
    $window_height = 0
    origin = Geom::Point3d.new(0,$width-$window_offset,$window_height)
    v1 = Geom::Vector3d.new(3,0,0)
    v2 = Geom::Vector3d.new(0,-$window_width,0)
    entities.add_line(origin+v1,origin+v1+v2)
    entities.add_line(origin+v1+v1,origin)
    entities.add_line(origin+v2,origin+v2+v1+v1)
    entities.add_line(origin+v1+v1,origin+v1+v1+v2)
    entities.add_line(origin,origin+v2)

    entities.add_text ($window_length/12).to_i.to_s+" X "+($window_width/12).to_i.to_s+ " WIN", origin+Geom::Vector3d.new(10,-24,0)
  else
    $window_height = 0
    origin = Geom::Point3d.new(0,$width-$window_offset,$window_height)
    v1 = Geom::Vector3d.new(3,0,0)
  	v2 = Geom::Vector3d.new(0,-$window_width,0)
    entities.add_line(origin+v1,origin+v1+v2)
  	entities.add_line(origin+v1+v1,origin)
  	entities.add_line(origin+v2,origin+v2+v1+v1)
  	entities.add_line(origin+v1+v1,origin+v1+v1+v2)
  	
  	entities.add_text ($window_length/12).to_i.to_s+" X "+($window_width/12).to_i.to_s+ " WIN", origin+Geom::Vector3d.new(10,-24,0)
  end

=begin
	origin = Geom::Point3d.new(0,$width-$window_offset,$window_height)
	v1 = Geom::Vector3d.new(0,0,$window_length)
	v2 = Geom::Vector3d.new(0,-$window_width,0)
	p1 = origin + v1
	p2 = origin + v1 + v2
	p3 = origin + v2
	a = entities.add_face origin, p1, p2, p3
	a.back_material = $window_color
	a.material = $window_color


	origin = origin + Geom::Vector3d.new(0,-2,2)
	v1 = Geom::Vector3d.new(0,0, $window_length -4)
	v2 = Geom::Vector3d.new(0, -$window_width + 4,0) 
	p1 = origin + v1
	p2 = origin + v1 + v2
	p3 = origin + v2
	b = entities.add_face origin, p1, p2, p3
	b.material = "[Translucent_Glass_Blue]"
	b.back_material = "[Translucent_Glass_Blue]"

	
	window_array = [a,b]


	if($window_type == "Verticle")
		p = origin
		
		origin = origin + Geom::Vector3d.new(-0.1,-0.1,0)
		v1 = Geom::Vector3d.new(0,0,($window_length-4)/2)
		v2 = Geom::Vector3d.new(0,-$window_width+4,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		c = entities.add_face origin, p1, p2, p3
		c.back_material = $window_color
		window_array.push(c)
		

		origin = origin + Geom::Vector3d.new(0,-2,2)
		v1 = Geom::Vector3d.new(0,0,(($window_length-4)/2) - 4)
		v2 = Geom::Vector3d.new(0,-$window_width+4+4,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		d = entities.add_face origin, p1, p2, p3
		d.back_material = "[Translucent_Glass_Blue]"
		d.material = "[Translucent_Glass_Blue]"
		window_array.push(d)
		
		
		
		
		origin = p + Geom::Vector3d.new(-1.4,0,($window_length-4)/2)
		v1 = Geom::Vector3d.new(0,0,($window_length-4)/2)
		v2 = Geom::Vector3d.new(0,-$window_width+4,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		e = entities.add_face origin, p1, p2, p3
		e.back_material = $window_color
		window_array.push(e)

		origin = origin + Geom::Vector3d.new(0,-2,2)
		v1 = Geom::Vector3d.new(0,0,(($window_length-4)/2) - 4)
		v2 = Geom::Vector3d.new(0,-$window_width+4+4,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		f = entities.add_face origin, p1, p2, p3
		f.back_material = "[Translucent_Glass_Blue]"
		f.material = "[Translucent_Glass_Blue]"
		window_array.push(f)
	end
	if($window_type == "Slider")
		p = origin
		
		origin = origin + Geom::Vector3d.new(-0.1,0,0)
		v1 = Geom::Vector3d.new(0,0,$window_length-4)
		v2 = Geom::Vector3d.new(0, -($window_width-4)/2,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		c = entities.add_face origin, p1, p2, p3
		c.back_material = $window_color
		window_array.push(c)

		origin = origin + Geom::Vector3d.new(0,-2,2)
		v1 = Geom::Vector3d.new(0,0,$window_length-4 - 4)
		v2 = Geom::Vector3d.new(0, -($window_width-4)/2+4,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		d = entities.add_face origin, p1, p2, p3
		d.back_material = "[Translucent_Glass_Blue]"
		d.material = "[Translucent_Glass_Blue]"
		window_array.push(d)
		
		
		origin = p + Geom::Vector3d.new(-1.4, -($window_width-4)/2,0)
		v1 = Geom::Vector3d.new(0,0,$window_length-4)
		v2 = Geom::Vector3d.new(0, -($window_width-4)/2,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		e = entities.add_face origin, p1, p2, p3
		e.back_material = $window_color
		window_array.push(e)
		
		
		origin = origin + Geom::Vector3d.new(0,-2,2)
		v1 = Geom::Vector3d.new(0,0,$window_length-4 - 4)
		v2 = Geom::Vector3d.new(0, -($window_width-4)/2+4,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		f = entities.add_face origin, p1, p2, p3
		f.back_material = "[Translucent_Glass_Blue]"
		f.material = "[Translucent_Glass_Blue]"
		window_array.push(f)
		
	end
	
	if($window_grid == 1)
		origin = Geom::Point3d.new(0,$width-$window_offset,$window_height) + Geom::Vector3d.new(0.1,-2,2)
		window_array.push(entities.add_line(origin + Geom::Vector3d.new(0,(-$window_width +4)/4,0), origin + Geom::Vector3d.new(0,(-$window_width +4)/4,0) + Geom::Vector3d.new(0,0, $window_length -4)))
		window_array.push(entities.add_line(origin + Geom::Vector3d.new(0,(-$window_width +4)*3/4,0), origin + Geom::Vector3d.new(0,(-$window_width +4)*3/4,0) + Geom::Vector3d.new(0,0, $window_length -4)))
		window_array.push(entities.add_line(origin + Geom::Vector3d.new(0,0,($window_length-4)/4), origin + Geom::Vector3d.new(0,0,($window_length-4)/4) +  Geom::Vector3d.new(0,-$window_width +4,0)))
		window_array.push(entities.add_line(origin + Geom::Vector3d.new(0,0,($window_length-4)*3/4), origin + Geom::Vector3d.new(0,0,($window_length-4)*3/4) +  Geom::Vector3d.new(0, -$window_width +4,0)))
	end
	
		
	if($window_shutter == 1)
		vs = Geom::Vector3d.new(0,14,0)
		v1 = Geom::Vector3d.new(0,0,$window_length)
		v2 = Geom::Vector3d.new(0,-$window_width,0)
		
		o = Geom::Point3d.new(-1,$width-$window_offset,$window_height)
		shutter1 = entities.add_face o, o+v1, o+v1+vs, o+vs
		shutter2 = entities.add_face o+v2, o+v2+v1,o+v2+v1-vs,o+v2-vs
		
		shutter1.material = "Kynar Matte Black Trim"
		shutter2.back_material = "Kynar Matte Black Trim"
	end
	
	
	window_group = entities.add_group window_array
	a.pushpull 2
	if($window_type != "Fixed")
		c.pushpull -1.2
		e.pushpull -1.2
	end
=end
end
end


def build_slide_door(entities,r,t)
#  define same variables



 posX = 0
 posY = 0
 posZ = 0

tracklen = (2*$slide_width)-12

case $slide_type
when "split"
   #1 split door
   trackstart = ((($slide_width/2)-6) *-1)
when "single left"
   #2  single   track to the left.  
   trackstart = (($slide_width)*(-1)+12)
else
   #3  single   track to the right
   trackstart = 0
end

    model = Sketchup.active_model

    model.start_operation "Create slidedoor"
 
    entities = model.active_entities

#----- comment lines if you dont want a group
   group = entities.add_group
   entities = group.entities

   entities.add_text(($slide_width/12).to_i.to_s+"X"+($slide_height/12).to_i.to_s + " SLD", [$slide_width/2,10,-2])

     pts = []
     pts[0] = [posX, posY, posZ]
     pts[1] = [(posX+$slide_width), posY, posZ]
     pts[2] = [(posX+$slide_width), posY, (posZ+$slide_height)]
     pts[3] = [posX, posY, (posZ+$slide_height)]
   base = entities.add_face pts
   base.material = $wallColor
    $slide_thickness = -$slide_thickness if( base.normal.dot(Z_AXIS) < 0 )
# Now we can do the pushpull
    base.pushpull $slide_thickness

	entities.add_line([0,-$slide_thickness,$slide_height],[0,0,$slide_height]).faces[1].material = $slide_frameColor
	entities.add_line([0,-$slide_thickness,$slide_height],[0,0,$slide_height]).faces[0].material = $slide_frameColor
	entities.add_line([$slide_width,-$slide_thickness,$slide_height],[$slide_width,0,$slide_height]).faces[0].material = $slide_frameColor
	
	p1 = Geom::Point3d.new(0,-$slide_thickness,0)
	p2 = Geom::Point3d.new(0,-$slide_thickness,$slide_height)
	p3 = Geom::Point3d.new($slide_width,-$slide_thickness,$slide_height)
	p4 = Geom::Point3d.new($slide_width,-$slide_thickness,0)
	v1 = Geom::Vector3d.new(-2,0,0)
	v2 = Geom::Vector3d.new(2,0,0)
	v3 = Geom::Vector3d.new(0,0,-2)
	slide_frame = entities.add_face(p1,p2,p3,p4,p4+v1,p3+v1+v3,p2+v2+v3,p1+v2).material = $slide_frameColor

 pts = []
 pts[0] = [posX+trackstart, posY, (posZ+$slide_height+3)]
 pts[1] = [posX+trackstart, (posY-6), (posZ+$slide_height+3)]
 pts[2] = [posX+trackstart, (posY-6), (posZ+$slide_height+5)]
 pts[3] = [posX+trackstart, (posY-2), (posZ+$slide_height+8)]
 pts[4] = [posX+trackstart, posY, (posZ+$slide_height+8)]

 # Add the face to the entities in the model
 face = entities.add_face pts
 face.back_material  = $track_color
#  pull track
 face.pushpull -tracklen




#  draw vert lines if a split door
if($slide_wainscot>0.1)
#   draw lines and add them to group
	entities.add_line([2,-$slide_thickness,$slide_wainscot],[$slide_width-2,-$slide_thickness,$slide_wainscot]).faces[1].material = $slide_wainscotcolor

=begin
 point1 = Geom::Point3d.new(posX, posY, (posZ+$slide_wainscot))
 point2 = Geom::Point3d.new(posX, (posY-$slide_thickness), (posZ+$slide_wainscot))
 line = entities.add_line point1,point2
 line.find_faces
 line.faces[0].material = $slide_color
 line.faces[1].material = $slide_wainscotcolor

 point1 = Geom::Point3d.new(posX, (posY-$slide_thickness), (posZ+$slide_wainscot))
 point2 = Geom::Point3d.new((posX+$slide_width), (posY-$slide_thickness), (posZ+$slide_wainscot))
 line = entities.add_line point1,point2
 line.find_faces
 line.faces[0].material = $slide_color
 line.faces[1].material = $slide_wainscotcolor
 
 point1 = Geom::Point3d.new((posX+$slide_width), (posY-$slide_thickness), (posZ+$slide_wainscot))
 point2 = Geom::Point3d.new((posX+$slide_width), posY, (posZ+$slide_wainscot))
 line = entities.add_line point1,point2
 line.find_faces
 line.faces[0].material = $slide_color
 line.faces[1].material = $slide_wainscotcolor
=end
end



case $slide_type
when "split"

#   draw lines and add them to group
 point1 = Geom::Point3d.new((posX+($slide_width/2)), posY, posZ)
 point2 = Geom::Point3d.new((posX+($slide_width/2)), (posY-$slide_thickness), posZ)
 line = entities.add_line point1,point2

 point1 = Geom::Point3d.new((posX+($slide_width/2)), (posY-$slide_thickness), posZ)
 point2 = Geom::Point3d.new((posX+($slide_width/2)), (posY-$slide_thickness), (posZ+$slide_height))
 line = entities.add_line point1,point2

 point1 = Geom::Point3d.new((posX+($slide_width/2)), (posY-$slide_thickness), (posZ+$slide_height))
 point2 = Geom::Point3d.new((posX+($slide_width/2)), posY, (posZ+$slide_height))
 line = entities.add_line point1,point2
end


  
#  draw wc lines

	#r = Geom::Transformation.rotation [0,0,0], [0,0,1],180.degrees
	entities.transform_entities r, group
	#t = Geom::Transformation.new [50,0,0]
	entities.transform_entities t, group


end





def get_ew1_points(posX, posY, posZ)
    pts = []
	pts.push([posX, posY, posZ])
	pts.push([posX, (posY + $width), posZ])
	if($oh2>0)
		pts.push([posX, (posY + $width), (posZ + $height2)])
	end
	pts.push([posX, (posY + $width + $oh2), (posZ + $height2)])
	pts.push([posX, (posY + ($width/2)), (posZ + $height2 + $gableht2)])
	
	if($oh1>0)
		  pts.push([posX, (posY - $oh1), (posZ + $height1)])
	end
	pts.push([posX, posY, (posZ + $height1)])
    return pts
  end
def build_porch(r, t, entities)
  brian = $roofColor
  $roofColor = $porch_roof_color
  h = $porch_length*$porch_pitch/12

  if($porch_type == "Hip")
	asmaa = $porch_length
  else
	asmaa = 0
  end
  
  x1 = 6
  y1 = 6
  x = 6
  y = 6
  c = (8.5-6)/2
	pt1 = Geom::Point3d.new(0,0,0)
	pt2 = Geom::Point3d.new($porch_width, 0, 0)
	pt3 = Geom::Point3d.new($porch_width, $porch_length, 0)
	pt4 = Geom::Point3d.new(0, $porch_length, 0)
	pt5 = Geom::Point3d.new(asmaa,0,0)
	pt6 = Geom::Point3d.new($porch_width-asmaa,0,0)
	v1 = Geom::Vector3d.new(0,0,-5.5)
	
	
	
	v = Geom::Vector3d.new(0,0, -$porch_height)
	porch_group = entities.add_group
	
	
	#porch_group.entities.add_cline pt1,pt2
	#face = porch_group.entities.add_face(pt1,pt2,pt3,pt4)
	#face.material = $faciaColor 
	#face.pushpull 5.5
	
	#porch_group.entities.add_face(pt1+v1,pt2+v1,pt3+v1,pt4+v1).material = $porch_ceiling_color
	
	
	
	porch_group.entities.add_cline pt2, pt3
	
	porch_group.entities.add_cline pt3, pt4
	porch_group.entities.add_cline(pt4,pt1)
	porch_group.entities.add_cline(pt3, pt6)#.find_faces
    porch_group.entities.add_cline(pt4, pt5)#.find_faces
	
	porch_group.entities.add_dimension_linear([$porch_overjet,$porch_length-$porch_overhang,0],[$porch_width-$porch_overjet,$porch_length-$porch_overhang,0], [0,60,0]).arrow_type = Sketchup::Dimension::ARROW_SLASH
	porch_group.entities.add_dimension_linear([$porch_overjet,0,0],[$porch_overhang,$porch_length-$porch_overhang,0], [-30,0,0]).arrow_type = Sketchup::Dimension::ARROW_SLASH
	
	#if(asmaa>0)
	#	porch_group.entities.add_face(pt1,pt4,pt5).material = $roofColor
	#	porch_group.entities.add_face(pt2,pt6,pt3).material = $roofColor
	#else
	#	porch_group.entities.add_face(pt1,pt4,pt5).material = $wallColor
	#	porch_group.entities.add_face(pt2,pt6,pt3).material = $wallColor
	#end
	#porch_group.entities.add_face(pt6,pt5,pt4,pt3).material = $roofColor

	if($post_number>0)
		a = porch_group.entities.add_face([$porch_overjet,0,0],[$porch_overjet,$porch_length-$porch_overhang,0],[$porch_width-$porch_overjet,$porch_length-$porch_overhang,0],[$porch_width-$porch_overjet,0,0], [$porch_width-$porch_overjet-8.5,0,0], [$porch_width-$porch_overjet-8.5,$porch_length-$porch_overhang-8.5,0],[$porch_overjet+8.5,$porch_length-$porch_overhang-8.5, 0], [$porch_overjet+8.5,0,0])
		v5 = Geom::Vector3d.new(0,($porch_width-2*$porch_overhang)/$post_number,0)
		pt = Geom::Point3d.new($porch_overjet,$porch_length-$porch_overhang,0)
		s = ($porch_width-$porch_overjet*2)/($post_number-1)
		for i in 0..($post_number-2)
			porch_group.entities.add_dimension_linear([$porch_overjet+s*i,$porch_length-$porch_overhang,0],[$porch_overjet+s*(i+1),$porch_length-$porch_overhang,0], [0,30,0]).arrow_type = Sketchup::Dimension::ARROW_SLASH
		end
		
		
		#a.material = $porch_header_color
		a.pushpull -5.5
		a.erase!
		a = porch_group.entities.add_line([$porch_overjet,0,0],[$porch_overjet+8.5,0,0])
		a.faces[0].erase!
		a.erase!
		a = porch_group.entities.add_line [$porch_width-$porch_overjet,0,0], [$porch_width-$porch_overjet-8.5,0,0]
		a.faces[0].erase!
		a.erase!
		a = porch_group.entities.add_face([$porch_overjet+c,$porch_length-$porch_overhang-c,0],[$porch_overjet+c+6,$porch_length-$porch_overhang-c,0],[$porch_overjet+c+6,$porch_length-$porch_overhang-c-6,0],[$porch_overjet+c,$porch_length-$porch_overhang-c-6,0])
		#a.material = $porch_post_color
		#a.pushpull $porch_height-11
		a = porch_group.entities.add_face([$porch_width-c-$porch_overjet,$porch_length-c-$porch_overhang,0],[$porch_width-c-$porch_overjet-6,$porch_length-$porch_overhang-c,0],[$porch_width-c-$porch_overjet-6,$porch_length-$porch_overhang-c-6,0],[$porch_width-$porch_overjet-c,$porch_length-$porch_overhang-c-6,0])
		#a.material = $porch_post_color
		#a.pushpull $porch_height-11
		if($post_number>2)
			d = ($porch_width-(2*c)-(2*$porch_overjet)-($post_number*6))/($post_number-1)
			for i in 1..($post_number-2)     
				a = porch_group.entities.add_face([$porch_overjet+c+i*(d+6),$porch_length-$porch_overhang-c,0],[$porch_overjet+c+6+i*(d+6),$porch_length-$porch_overhang-c,0],[$porch_overjet+c+6+i*(d+6),$porch_length-$porch_overhang-c-6,0],[$porch_overjet+c+i*(d+6),$porch_length-$porch_overhang-c-6,0])
				#a.material = $porch_post_color
				#a.pushpull $porch_height-11
			end
		end
	end
	entities.transform_entities r, porch_group
	entities.transform_entities t, porch_group	
	$roofColor = brian
end 
def create_porch(entities)
	$porch_height = -4
	if($porch_side == "SW1")
		r = Geom::Transformation.rotation [$porch_width/2, 0, 0], [0,0,1], 180.degrees
		t = Geom::Transformation.new [$porch_offset,0, $porch_height]
		build_porch(r, t, entities)
	end
  
	if($porch_side == "SW2")
		r = Geom::Transformation.rotation [$porch_width/2, 0, 0], [0,0,1], 360.degrees
		t = Geom::Transformation.new [$length-$porch_offset-$porch_width,$width, $porch_height]
		build_porch(r, t, entities)
	end
  
	if($porch_side == "EW1")
		r = Geom::Transformation.rotation [0, 0, 0], [0,0,1], 90.degrees
		t = Geom::Transformation.new [0,$width-$porch_offset-$porch_width, $porch_height]
		build_porch(r, t, entities)
	end
  
	if($porch_side == "EW2")
		r = Geom::Transformation.rotation [0, 0, 0], [0,0,1], 270.degrees
		t = Geom::Transformation.new [$length, $porch_offset+$porch_width, $porch_height]
		build_porch(r, t, entities)
	end
end

$size1 = worksheet.Cells(114,1).Value.to_i
for i in 1..$size1
	$concrete_width = worksheet.Cells(115,1+i).Value.to_f
	$concrete_length = worksheet.Cells(116,1+i).Value.to_f
	$concrete_side = worksheet.Cells(117,1+i).Value
	$concrete_offset = worksheet.Cells(118,1+i).Value.to_f
	build_concrete(entities)
end



# concrete_data.each do |concrete|
	# $concrete_width = concrete[0]
	# $concrete_length = concrete[1]
	# $concrete_side = concrete[2]
	# $concrete_offset = concrete[3]
	# build_concrete(entities)
# end



#porch_excel
$size1 = 12
if($size1>0)
	for i in 1..$size1

		$porch_offset = worksheet.Cells(63,1+i).Value.to_f

		$porch_width = worksheet.Cells(64,1+i).Value.to_f

		$porch_length = worksheet.Cells(65,1+i).Value.to_f

		$porch_side = worksheet.Cells(66,1+i).Value
		$post_number = worksheet.Cells(67,1+i).Value.to_f
		$porch_type = worksheet.Cells(68,1+i).Value

		$porch_pitch = worksheet.Cells(69,1+i).Value.to_f
		$porch_overhang = worksheet.Cells(70,1+i).Value.to_f
		
    $porch_overjet = $porch_overhang

		#$porch_width = $porch_width+$porch_overhang*2
		create_porch(entities)
	end
end



# porch_data.each do |porch|
	# $porch_offset = porch[0]
	# $porch_width = porch[1]
	# $porch_length = porch[2]
	# $porch_side = porch[3]
	# $post_number = porch[4]
	# $porch_type = porch[5]
	# create_porch(entities)
# end

# post_data.each do |post|
	# $post_side = post[0]
	# $post_offset = post[1]
	# $post_width = post[2]
	# $post_length = post[3]
	# build_post(entities)
# end







$size1 = 100


def create_cupola(entities)

	if($cupola_number == 1)
		t = Geom::Transformation.new [$length/2,$width/2 - $cub_size/2, -10]
		build_cupola($cub_size, entities, $pitch, t, $roofColor, $wallColor)
	end
  
	if($cupola_number == 2)
		
	  t = Geom::Transformation.new [$length/4,$width/2 - $cub_size/2, -10]
	  build_cupola($cub_size, entities, $pitch, t, $roofColor, $wallColor)
	  t = Geom::Transformation.new [$length*3/4,$width/2 - $cub_size/2, -10] 
	  build_cupola($cub_size, entities, $pitch, t, $roofColor, $wallColor)
	end
	  
	if($cupola_number == 3)
		t = Geom::Transformation.new [$length/6,$width/2 - $cub_size/2, -10]
		build_cupola($cub_size, entities, $pitch, t,$roofColor, $wallColor)
		t = Geom::Transformation.new [($length/2)-$cub_mid/2,$width/2 - $cub_mid/2, -10]
		build_cupola($cub_mid, entities, $pitch, t,$roofColor, $wallColor)
		t = Geom::Transformation.new [$length*5/6-$cub_size,$width/2 - $cub_size/2, -10]
		build_cupola($cub_size, entities, $pitch, t,$roofColor, $wallColor)
	end
end
def create_sidelight(entities)
	if ($sidelight_side == "SW1")
		origin = Geom::Point3d.new($corner, 0, $height1)
		v1 = Geom::Vector3d.new($length-2*$corner,0,0)
		v2 = Geom::Vector3d.new(0,0,$sidelight_down)
		pt1 = origin + v1
		pt2 = origin + v1 - v2
		pt3 = origin - v2
		a = entities.add_face(origin,pt1,pt2,pt3)
		a.erase!
		b = entities.add_face origin,pt1,pt2,pt3
		b.material = "[Translucent_Glass_Blue]"
		b.back_material = "[Translucent_Glass_Blue]"
	end

	if ($sidelight_side == "SW2")
		origin = Geom::Point3d.new($corner, $width, $height2)
		v1 = Geom::Vector3d.new($length-2*$corner,0,0)
		v2 = Geom::Vector3d.new(0,0,$sidelight_down)
		pt1 = origin + v1
		pt2 = origin + v1 - v2
		pt3 = origin - v2
		a = entities.add_face origin,pt1,pt2,pt3
		a.erase!
		a = entities.add_face origin,pt1,pt2,pt3
		a.material = "[Translucent_Glass_Blue]"
		a.back_material = "[Translucent_Glass_Blue]"
	end
end


def create_slide_door(entities)
  if($slide_side == "EW2")
	r = Geom::Transformation.rotation [0, 0, 0], [0,0,1], 90.degrees
	t = Geom::Transformation.new [$length, $slide_offset,0]
	build_slide_door(entities, r, t)
  end
  
  if($slide_side == "EW1")
	r = Geom::Transformation.rotation [0, 0, 0], [0,0,1], 270.degrees
	t = Geom::Transformation.new [0, $width - $slide_offset,0]
	build_slide_door(entities, r, t)
  end
  
  if($slide_side == "SW1")
	r = Geom::Transformation.rotation [0, 0, 0], [0,0,1], 360.degrees
	t = Geom::Transformation.new [$slide_offset, 0,0]
	build_slide_door(entities, r, t)
  end
  
  if($slide_side == "SW2")
	r = Geom::Transformation.rotation [0, 0, 0], [0,0,1], 180.degrees
	t = Geom::Transformation.new [$length-$slide_offset, $width,0]
	build_slide_door(entities, r, t)
  end
end


create_interior(entities)
  


create_cupola(entities)





	
# window_data.each do |window|
	# $window_side = window[0]
	# $window_offset = window[1]
	# $window_height = window[2]
	# $window_length = window[3]
	# $window_width = window[4]
	# $window_type = window[5]
	# create_window(entities)
 # end






$LIS = false

def build_basic(entities, model)


if($LIS)
	entities.add_line([2+$post_length,0,0],[2+$post_length,$width,0])
	entities.add_line([$length-2-$post_length,0,0],[$length-2-$post_length,$width,0])
	entities.add_line([2+$post_length,2+$post_length,0],[$length-2-$post_length,2+$post_length,0])
	entities.add_line([2+$post_length,$width-2-$post_length,0],[$length-2-$post_length,$width-2-$post_length,0])
end

ew1 = entities.add_face get_ew1_points($posX, $posY, $posZ)
ew1.material = $wallColor
ew1.reverse!

logo_z_value = $height1 + ($gableht1 / 2) 
logo_y_value = (($width + 15) / 2) 
#entities.add_face([1,1,0],[1,$width-1,0],[$length-1,$width-1,0],[$length-1,1,0]).material = "white"
 if (has_apron?)
    # Create a series of "points", each a 3-item array containing x, y, and z.
    pt1 = [($posX - $apron), ($posY - $apron), $posZ]
    pt2 = [($posX + $length + $apron), ($posY - $apron), $posZ]
    pt3 = [($posX + $length + $apron), ($posY + $width + $apron), $posZ]
    pt4 = [($posX - $apron), ($posY + $width + $apron), $posZ]
    new_face = entities.add_face pt1, pt2, pt3, pt4
	materials = model.materials
	# m1 = materials.add('Vegetation_Blur7')
	# save_path = Sketchup.find_support_file "Vegetation/grass.jpg", "Materials"
	# m1.texture = save_path
	new_face.back_material = $grass_color
  end
  
  
  
  
  add_3d_letter_ew1 entities, "G", logo_y_value, logo_z_value
  add_3d_letter_ew2 entities, "B", logo_y_value -7.5, logo_z_value, $length

  add_3d_letter_ew1 entities, "B", logo_y_value - 7.5, logo_z_value - 7.5
  add_3d_letter_ew2 entities, "G", logo_y_value - 15, logo_z_value + 7.5, $length



  ew2 = entities.add_face get_ew2_points($posX, $posY, $posZ)
  ew2.material = $wallColor
  

  sw1 = entities.add_face get_sw1_points($posX, $posY, $posZ)
  sw1.material = $wallColor

  sw2 = entities.add_face get_sw2_points($posX, $posY, $posZ)
  sw2.material = $wallColor

  faciaEW1 = entities.add_face get_facia_ew1_points($posX, $posY, $posZ)
  faciaEW1.material = $faciaColor

  faciaEW2 = entities.add_face get_facia_ew2_points($posX, $posY, $posZ)
  faciaEW2.material = $faciaColor
  faciaEW2.back_material = $faciaColor

  
   if (has_overhang1?)         
    sofitSW1 = entities.add_face get_oh_sofit_sw1_points($posX, $posY, $posZ)
    sofitSW1.material = $sofitColor
   end
   
   if(has_overhang2?)
    sofitSW2 = entities.add_face get_oh_sofit_sw2_points($posX, $posY, $posZ)
    sofitSW2.material = $sofitColor
  end

  # draw oj sofits. 
  if (has_overjet1?)    
    sofitEW1right = entities.add_face get_oj_sofit_ew1_right_points($posX, $posY, $posZ)
    sofitEW1right.material = $sofitColor
    
    sofitEW1left = entities.add_face get_oj_sofit_ew1_left_points($posX, $posY, $posZ)
    sofitEW1left.material = $sofitColor
  end
  
  if(has_overjet2?)
    sofitEW2left = entities.add_face get_oj_sofit_ew2_left_points($posX, $posY, $posZ)
    sofitEW2left.material = $sofitColor
    
    sofitEW2right = entities.add_face get_oj_sofit_ew2_right_points($posX, $posY, $posZ)
    sofitEW2right.material = $sofitColor
  end

  ### draw facias end walls 
  faciaEW1right = entities.add_face get_facia_ew1_right_points($posX, $posY, $posZ)
  faciaEW1right.material = $faciaColor
  
  faciaEW1left = entities.add_face get_facia_ew1_left_points($posX, $posY, $posZ)
  faciaEW1left.material = $faciaColor

  # Add the face to the entities in the model
  faciaEW2right = entities.add_face get_facia_ew2_right_points($posX, $posY, $posZ)
  faciaEW2right.material = $faciaColor
  
  # Add the face to the entities in the model
  faciaEW2left = entities.add_face get_facia_ew2_left_points($posX, $posY, $posZ)
  faciaEW2left.material = $faciaColor

  ### Draw the roof.
  roofSW1 = entities.add_face get_roof_sw1_points($posX, $posY, $posZ)
  roofSW1.material = $roofColor

  roofSW2 = entities.add_face get_roof_sw2_points($posX, $posY, $posZ)
  roofSW2.material = $roofColor

  # drawn corner trim lines on endwalls based on OH
  if (has_overhang1? or has_overhang2?)
    # draw corner lines on rear end wall and horz line at top
    entities.add_line([$posX, ($posY + $corner), $posZ],[$posX, ($posY + $corner), ($posZ + $height1)])
    entities.add_line([$posX, ($posY + $corner), ($posZ + $height1)],[$posX, ($posY), ($posZ + $height1)])

    entities.add_line([$posX, ($posY + $width-$corner), $posZ],[$posX, ($posY + $width-$corner), ($posZ + $height2)])
    entities.add_line([$posX, ($posY + $width-$corner), ($posZ + $height2)],[$posX, ($posY + $width), ($posZ + $height2)])

    # draw $corner lines on front end wall and horz line at top
    entities.add_line([($posX + $length), ($posY + $corner), $posZ],[($posX + $length), ($posY + $corner), ($posZ + $height1)])
    entities.add_line([($posX + $length), ($posY + $corner), ($posZ + $height1)],[($posX + $length), ($posY), ($posZ + $height1)])

    entities.add_line([($posX + $length), ($posY + $width-$corner), $posZ],[($posX + $length), ($posY + $width-$corner), ($posZ + $height2)])
    entities.add_line([($posX + $length), ($posY + $width-$corner), ($posZ + $height2)],[($posX + $length), ($posY + $width), ($posZ + $height2)])

  else
    # no OH, draw taller $corner trim lines on endwalls

    # draw $corner lines on rear end wall
    entities.add_line([$posX, ($posY + $corner), $posZ],[$posX, ($posY + $corner), ($posZ + $height + ($pitch*$corner/12))])
    entities.add_line([$posX, ($posY + $width-$corner), $posZ],[$posX, ($posY + $width-$corner), ($posZ + $height + ($pitch*$corner/12))])

    # draw $corner lines on front end wall
    entities.add_line([($posX + $length), ($posY + $corner), $posZ],[($posX + $length), ($posY + $corner), ($posZ + $height1 + ($pitch*$corner/12))])
    entities.add_line([($posX + $length), ($posY + $width-$corner), $posZ],[($posX + $length), ($posY + $width-$corner), ($posZ + $height + ($pitch*$corner/12))])

  end

  # draw eave lines on side walls
  # when moved here ahead of side building.corner lines, the side wall section is separate of building.corner trim areas
  entities.add_line([$posX, $posY, ($posZ + $height1)],[($posX + $length), $posY, ($posZ + $height1)])
  entities.add_line([$posX, ($posY + $width), ($posZ + $height2)],[($posX + $length), ($posY + $width), ($posZ + $height2)])

  # draw the same ht lines on sides

  # draw $corner lines on back side wall (y=$width)
  entities.add_line([($posX + $corner), ($posY + $width), $posZ],[($posX + $corner), ($posY + $width), ($posZ + $height2)])
  entities.add_line([($posX + $length-$corner), ($posY + $width), $posZ],[($posX + $length-$corner), ($posY + $width), ($posZ + $height2)])

  # draw $corner lines on back side wall (y=0)
  entities.add_line([($posX + $corner), $posY, $posZ],[($posX + $corner), $posY, ($posZ + $height1)])
  entities.add_line([($posX + $length-$corner), $posY, $posZ],[($posX + $length-$corner), $posY, ($posZ + $height1)])

  # draw wainscot lines
  if (has_wainscot?)
    entities.add_line([($posX + $corner), ($posY + $width), ($posZ + $wcht)],[($posX + $length-$corner), ($posY + $width), ($posZ + $wcht)])
    entities.add_line([$posX, ($posY + $corner), ($posZ + $wcht)],[$posX, ($posY + $width-$corner), ($posZ + $wcht)])
    entities.add_line([($posX + $length), ($posY + $corner), ($posZ + $wcht)],[($posX + $length), ($posY + $width-$corner), ($posZ + $wcht)])
    
    ### Draw the wainscot
    wainscotSW1 = entities.add_face get_wainscot_sw1_points($posX, $posY, $posZ)
    wainscotSW1.material = $wainscotColor
	wainscotSW1.back_material = $wainscotColor

    wainscotSW2 = entities.add_face get_wainscot_sw2_points($posX, $posY, $posZ)
    wainscotSW2.material = $wainscotColor
	wainscotSW2.back_material = $wainscotColor
    
    wainscotEW1 = entities.add_face get_wainscot_ew1_points($posX, $posY, $posZ)
    wainscotEW1.material = $wainscotColor
    wainscotEW1.back_material = $wainscotColor

    wainscotEW2 = entities.add_face get_wainscot_ew2_points($posX, $posY, $posZ)
    wainscotEW2.material = $wainscotColor
	wainscotEW2.back_material = $wainscotColor

	a = entities.add_line([0,0,$wcht],[0,$corner,$wcht])
	a.faces[0].back_material = $wainscotColor
	entities.add_line([0,$corner,0],[0,$corner,$wcht]).erase!

	a = entities.add_line([0,0,$wcht],[$corner,0,$wcht])
	a.faces[1].back_material = $wainscotColor
	entities.add_line([$corner,0,0],[$corner,0,$wcht]).erase!

	a = entities.add_line([$length,0,$wcht],[$length-$corner,0,$wcht])
	a.faces[0].back_material = $wainscotColor
	entities.add_line([$length-$corner,0,0],[$length-$corner,0,$wcht]).erase!
	
	a = entities.add_line([$length,$corner,$wcht],[$length,0,$wcht])
	a.faces[0].back_material = $wainscotColor
	entities.add_line([$length,$corner,$wcht],[$length,$corner,0]).erase!
	
	a = entities.add_line([$length,$width,$wcht],[$length-$corner,$width,$wcht])
	a.faces[1].back_material = $wainscotColor
	entities.add_line([$length-$corner,$width,0],[$length-$corner,$width,$wcht]).erase!
	
	a = entities.add_line([$length,$width,$wcht],[$length,$width-$corner,$wcht])
	a.faces[0].back_material = $wainscotColor
	entities.add_line([$length,$width-$corner,$wcht],[$length,$width-$corner,0]).erase!
	
	a = entities.add_line([0,$width,$wcht],[0,$width-$corner,$wcht])
	a.faces[1].back_material = $wainscotColor
	entities.add_line([0,$width-$corner,0],[0,$width-$corner,$wcht]).erase!
	
	a = entities.add_line([0,$width,$wcht],[$corner,$width,$wcht])
	a.faces[0].back_material = $wainscotColor
	entities.add_line([$corner,$width,0],[$corner,$width,$wcht]).erase!
	
	
	if($concrete_height>0)
		a = entities.add_line([0,0,$concrete_height],[$length,0,$concrete_height])
		a.faces[1].back_material = $concrete_color
		a.faces[1].material = $concrete_color
		
		a = entities.add_line([$length,0, $concrete_height],[$length,$width,$concrete_height])
		a.faces[1].back_material = $concrete_color
		a.faces[1].material = $concrete_color
		
		a = entities.add_line([$length,$width,$concrete_height],[0,$width,$concrete_height])
		a.faces[1].back_material = $concrete_color
		a.faces[1].material = $concrete_color
		
		a = entities.add_line([0,$width,$concrete_height],[0,0,$concrete_height])
		a.faces[1].back_material = $concrete_color
		a.faces[1].material = $concrete_color	
	end
  end
  

	if($gambrel_height>0)
		top = Geom::Point3d.new(-$oj1,$width/2,$height1+$gableht1)
		v1 = Geom::Vector3d.new(0,0,-$brian*$pitch/12)
		v2 = Geom::Vector3d.new(0,$brian,0)
		v3 = Geom::Vector3d.new($length+$oj1+$oj2,0,0)
		v4 = Geom::Vector3d.new(0,0,+$facia)
		pt1 = top + v1 + v2
		pt2 = top + v1 + v2 + v3

		p3 = top + v1 - v2
		p4 = top + v1 - v2 + v3

		t = Geom::Transformation.new [0,0, $gambrel_height]

		line2 = entities.add_line([-$oj1,$width/2,$height1+$gableht1+$facia],[$length+$oj2, $width/2,$height1+$gableht1+$facia])
		line4 = entities.add_line(pt1+v4,pt2+v4)
		line6 = entities.add_line(p3+v4, p4+v4)
		entities.transform_entities t, line2



		if($oj1 == 0)
			$oj1 = 1
		end

			line111 = entities.add_line(top, top-Geom::Vector3d.new(-$oj1,0,0))
			line333 = entities.add_line(pt1,pt1-Geom::Vector3d.new(-$oj1,0,0))
			line555 = entities.add_line(p3,p3-Geom::Vector3d.new(-$oj1,0,0))

		if($oj2==0)
			$oj2 = 1
		end

			line11 = entities.add_line(top+v3, top+v3-Geom::Vector3d.new($oj2,0,0))
			line33 = entities.add_line(pt1+v3,pt1+v3-Geom::Vector3d.new($oj2,0,0))
			line55 = entities.add_line(p3+v3,p3+v3-Geom::Vector3d.new($oj2,0,0))


		entities.transform_entities t, line4
		entities.transform_entities t, line6

			entities.transform_entities t, line111
			entities.transform_entities t, line333
			entities.transform_entities t, line555



			entities.transform_entities t, line11
			entities.transform_entities t, line33
			entities.transform_entities t, line55

		end
end


$post_length = worksheet.Cells(75,2).Value.to_f
build_basic(entities, model)

# walkdoor_data.each do |walkdoor|
	# $walkside = walkdoor[0]
	# $walk_height = walkdoor[1]
	# $walk_width = walkdoor[2]
	# $walk_color = walkdoor[3]
	# $walk_offset = walkdoor[4]
	# $walk_window = walkdoor[5]
	# $walk_grid = walkdoor[6]
	# $door_swing = walkdoor[7]
	# create_walkdoor(entities)
# end 

$size1 = 20
if($size1>0)
	for i in 1..$size1
			$window_side = worksheet.Cells(107,1+i).Value
			$window_offset = worksheet.Cells(108,1+i).Value.to_f
			$window_height = worksheet.Cells(109,1+i).Value.to_f
			$window_length = worksheet.Cells(110,1+i).Value.to_f
			$window_width = worksheet.Cells(111,1+i).Value.to_f
			$window_type = worksheet.Cells(112,1+i).Value
			$window_color = worksheet.Cells(113,1+i).Value
			$window_grid = worksheet.Cells(114,1+i).Value.to_i
			$window_shutter = worksheet.Cells(106,1+i).Value.to_i
			create_window(entities)
	end
end

$size1 = 13
if($size1>0)
	for i in 1..$size1	
			$walkside = worksheet.Cells(32,1+i).Value
			$walk_height = worksheet.Cells(33,1+i).Value.to_f
			$walk_width = worksheet.Cells(34,1+i).Value.to_f
			$walk_color = worksheet.Cells(35,1+i).Value
			$walk_offset = worksheet.Cells(36,1+i).Value.to_f
			$walk_window = worksheet.Cells(37,1+i).Value
			$walk_grid = worksheet.Cells(38,1+i).Value
			$door_swing = worksheet.Cells(39,1+i).Value
			create_walkdoor(entities)
	end
end


#overhead_data
$size1 = 13
if($size1>0)
	for i in 1..$size1
    $offset_length = worksheet.Cells(24,1+i).Value.to_f
    $door_height = worksheet.Cells(25,1+i).Value.to_f
    $door_width = worksheet.Cells(26, 1+i).Value.to_f
    $panel = worksheet.Cells(27, 1+i).Value.to_f
    $side = worksheet.Cells(28, 1+i).Value
    $overheadColor = worksheet.Cells(29,1+i).Value
    $overhead_window = worksheet.Cells(151,1+i).Value.to_i
    $overhead_window_width = worksheet.Cells(152,1+i).Value.to_f
    $overhead_window_height = worksheet.Cells(153,1+i).Value.to_f
    $overhead_opening = worksheet.Cells(30,1+i).Value
    $dutch = worksheet.Cells(23,1+i).Value.to_f
    $wall_opening = worksheet.Cells(22,1+i).Value.to_f
		create_overhead(entities)
	end
end

$size1 = 100
$post_point_SW1 = []
$post_point_SW1_vector = []
$post_point_EW2 = []
$post_point_EW2_vector = []
$post_point_SW2 = []
$post_point_SW2_vector = []
$post_point_EW1 = []
$post_point_EW1_vector = []
def add_post_point
	c = $oh1+20
	if($post_side == "SW1")
		$post_point_SW1.push([$post_offset+$post_length/2,0,0])
		$post_point_SW1_vector.push([0,-c,0])
	end
	if($post_side == "SW2")
		$post_point_SW2.push([$length-$post_offset-$post_length/2,$width,0])
		$post_point_SW2_vector.push([0,c,0])
	end
	if($post_side == "EW2")	
		if($post_offset<3)
			$post_point_EW2.push([$length,0,0])
			$post_point_EW2_vector.push([c,0,0])
		elsif($width-$post_offset-$post_length<3)
			$post_point_EW2.push([$length,$width,0])
			$post_point_EW2_vector.push([c,0,0])
		else
			$post_point_EW2.push([$length,$post_offset+$post_length/2,0])
			$post_point_EW2_vector.push([c,0,0])
		end
	end
	if($post_side == "EW1")
		if($post_offset<3)
			$post_point_EW1.push([0,$width,0])
			$post_point_EW1_vector.push([-c,0,0])
		elsif($width-$post_offset-$post_length<3)
			$post_point_EW1.push([0,0,0])
			$post_point_EW1_vector.push([-c,0,0])
		else
			$post_point_EW1.push([0,$width-$post_offset-$post_length/2,0])
			$post_point_EW1_vector.push([-c,0,0])
		end
	end
end

if($size1>0)
	for i in 1..$size1
		$post_side = worksheet.Cells(72,1+i).Value
		$post_offset = worksheet.Cells(73,1+i).Value.to_f*12
		$post_width = worksheet.Cells(74,1+i).Value.to_f
		$post_length = worksheet.Cells(75,1+i).Value.to_f
		$post_x = worksheet.Cells(76,1+i).Value.to_f
		$circle = worksheet.Cells(128,i+1).Value.to_i
		#$square = '%.2f' % worksheet.Cells(129,1+i).Value.to_f
    $square = worksheet.Cells(129,1+i).Value.to_i
		if($post_side == "SW1" or $post_side =="SW2")
			$post_offset = $post_offset-$post_length/2
		end
		build_post(entities)
		add_post_point
	end
end






# overhead_data.each do |overhead|
	# UI.messagebox("irene")
	# $offset_length = overhead[0]
	# $door_height = overhead[1]
	# $door_width = overhead[2]
	# $panel = overhead[3]
	# $side = overhead[4]
	# $overheadColor = overhead[5]
# end

$size1 = 2
if($size1>0)
	for i in 1..$size1
		$sidelight_side = worksheet.Cells(87,1+i).Value
		$sidelight_down = worksheet.Cells(88,1+i).Value.to_f
		create_sidelight(entities)
	end
end
		#create_sidelight(entities)

$size1 = 10
if($size1>0)
	for i in 1..$size1
		$slide_width = worksheet.Cells(48,1+i).Value.to_f
		$slide_height = worksheet.Cells(49,1+i).Value.to_f
		$slide_thickness = 4
		$slide_wainscot = worksheet.Cells(50,1+i).Value.to_f
		$slide_type = worksheet.Cells(51, 1+i).Value
		$slide_side = worksheet.Cells(52, 1+i).Value
		$slide_offset = worksheet.Cells(53, 1+i).Value.to_f
		$slide_color = worksheet.Cells(54, 1+i).Value
		$slide_wainscotcolor = worksheet.Cells(55, 1+i).Value
		$track_color = worksheet.Cells(56,1+i).Value
		if($slide_offset > 0)
			draw_rectangle($slide_side, $slide_offset, $slide_height, $slide_width, entities, 0)
		end
	end
end



# slide_data.each do |slide|
	# $slide_width = slide[0]
	# $slide_height = slide[1]
	# $slide_thickness = slide[2]
	# $slide_wainscot = slide[3]
	# $slide_type = slide[4]
	# $slide_side = slide[5]
	# $slide_offset = slide[6]
	# $slide_color = slide[7]
	# $slide_wainscotcolor = slide[8]
	# $track_color = slide[9]
	# draw_rectangle($slide_side, $slide_offset, $slide_height, $slide_width, entities, 0)

# end
# hydraulic_data.each do |hydraulic|
	# $hydraulic_wainscot = hydraulic[0] 
	# $hydraulic_height = hydraulic[1]
	# $hydraulic_width = hydraulic[2]
	# $hydraulic_wainscot_color = hydraulic[3]
	# $hydraulic_color = hydraulic[4]
	# $hydraulic_side = hydraulic[5]
	# $hydraulic_offset = hydraulic[6]
	# draw_rectangle($hydraulic_side, $hydraulic_offset, $hydraulic_height+8+6, $hydraulic_width, entities, 0)
# end

$size1 = 10
if($size1>0)
	for i in 1..$size1
		$hydraulic_wainscot = worksheet.Cells(79,1+i).Value.to_f
		$hydraulic_height = worksheet.Cells(80,1+i).Value.to_f
		$hydraulic_width = worksheet.Cells(81,1+i).Value.to_f
		$hydraulic_wainscot_color = worksheet.Cells(82,1+i).Value
		$hydraulic_color = worksheet.Cells(83,1+i).Value
		$hydraulic_side = worksheet.Cells(84,1+i).Value
		$hydraulic_offset = worksheet.Cells(85,1+i).Value.to_f
		if($hydraulic_offset>0)
			draw_rectangle($hydraulic_side, $hydraulic_offset, $hydraulic_height+8+6, $hydraulic_width, entities, 0)
		end
	end
end

$size1 = 10
if($size1>0)
	for i in 1..$size1
		$slide_width = worksheet.Cells(48,1+i).Value.to_f
		$slide_height = worksheet.Cells(49,1+i).Value.to_f
		$slide_thickness = 4
		$slide_wainscot = worksheet.Cells(50,1+i).Value.to_f
		$slide_type = worksheet.Cells(51, 1+i).Value
		$slide_side = worksheet.Cells(52, 1+i).Value
		$slide_offset = worksheet.Cells(53, 1+i).Value.to_f
		$slide_color = worksheet.Cells(54, 1+i).Value
		$slide_wainscotcolor = worksheet.Cells(55, 1+i).Value
		$track_color = worksheet.Cells(56,1+i).Value
		if($slide_offset > 0)
			create_slide_door(entities)
		end
	end
end


# slide_data.each do |slide|
	# $slide_width = slide[0]
	# $slide_height = slide[1]
	# $slide_thickness = slide[2]
	# $slide_wainscot = slide[3]
	# $slide_type = slide[4]
	# $slide_side = slide[5]
	# $slide_offset = slide[6]
	# $slide_color = slide[7]
	# $slide_wainscotcolor = slide[8]
	# $track_color = slide[9]
	# create_slide_door(entities)
# end





def interior_steel(entities)
#draw_rectangle("EW1",70*12, 14*12+8, 20*12, entities)

	if($interior_steel != "")
	
		if($oh1 >= $oh2)
		a = entities.add_line([0,0,$height],[0,$width,$height])
		b = entities.add_line([0,0,$height],[0,0,$height1])
					a.faces[0].back_material = $interior_steel
		# if($oh1 <= $oh2)

		# else
			# a.faces[1].back_material = $interior_steel
		# end
		
		#entities.add_line([$length,0,$height],[$length,0,$height1])


		f = entities.add_line([$length,0,$height],[$length,0,$height1])
		e = entities.add_line([$length,0,$height],[$length,$width,$height])
		e.faces[1].back_material = $interior_steel

		entities.add_line([0,0,$height],[$length,0,$height])
		d = entities.add_line([$corner,0,$height1],[$length-$corner, 0, $height1])
		
		
		
		d.find_faces
		d.faces[0].back_material = $interior_steel
		d.faces[2].back_material = $interior_steel
		if($height1 == $height2)
			d.faces[2].erase!
		end

		c = entities.add_line([$corner,$width,$height2],[$length-$corner,$width,$height2])
		c.faces[0].back_material = $interior_steel
		end
		
		if($oh1 < $oh2)
		a = entities.add_line([0,0,$height],[0,$width,$height])
		b = entities.add_line([0,$width,$height],[0,$width,$height2])
					a.faces[0].back_material = $interior_steel
		# if($oh1 <= $oh2)

		# else
			# a.faces[1].back_material = $interior_steel
		# end
		
		#entities.add_line([$length,0,$height],[$length,0,$height1])


		f = entities.add_line([$length,$width,$height],[$length,$width,$height2])
		e = entities.add_line([$length,0,$height],[$length,$width,$height])
		e.faces[1].back_material = $interior_steel

		entities.add_line([0,0,$height],[$length,0,$height])
		d = entities.add_line([$corner,0,$height1],[$length-$corner, 0, $height1])
		
		
		
		d.find_faces
		d.faces[1].back_material = $interior_steel
				d.faces[0].back_material = $interior_steel
		#d.faces[2].back_material = $interior_steel
		if($height1 == $height2)
			d.faces[2].erase!
		end

		c = entities.add_line([$corner,$width,$height2],[$length-$corner,$width,$height2])
		c.faces[0].back_material = $interior_steel
		end
		
		
		
	end
end

unless $interior_steel == "None"
	interior_steel(entities)
end


#always at bottom
if($wcht>0 && $concrete_height == 0)
	a = entities.add_line([0,5.5,0],[0,5.5,$wcht])
	a.faces[1].material = $wainscot_corner
	a = entities.add_line([5.5,0,0],[5.5,0,$wcht])
	a.faces[0].material = $wainscot_corner
	a = entities.add_line([$length-5.5,0,0],[$length-5.5,0,$wcht])
	a.faces[1].material = $wainscot_corner
	a = entities.add_line([$length,5.5,0],[$length,5.5,$wcht])
	a.faces[0].material = $wainscot_corner
	a = entities.add_line([$length,$width-5.5,0],[$length,$width-5.5,$wcht])
	a.faces[1].material = $wainscot_corner
	a = entities.add_line([$length-5.5,$width,0],[$length-5.5,$width,$wcht])
	a.faces[0].material = $wainscot_corner
	a = entities.add_line([0,$width-5.5,0],[0,$width-5.5,$wcht])
	a.faces[0].material = $wainscot_corner
	a = entities.add_line([5.5,$width,0],[5.5,$width,$wcht])
	a.faces[1].material = $wainscot_corner
end


a = entities.add_line([0,0,$wcht],[0,0,$height1])
a.faces[1].material = $main_color
a.faces[0].material = $main_color
a = entities.add_line([$length,0,$wcht],[$length,0,$height1])
a.faces[1].material = $main_color
a.faces[0].material = $main_color
a = entities.add_line([$length,$width,$wcht],[$length,$width,$height2])
a.faces[1].material = $main_color
a.faces[0].material = $main_color
a = entities.add_line([0,$width,$wcht],[0,$width,$height2])
a.faces[1].material = $main_color
a.faces[0].material = $main_color



if($gambrel_height==0)
	if($ridge == "[Translucent_Glass_Blue]")
		b = 24
	else	
		b = 5
	end
kathy = Geom::Point3d.new(-$oj1,($width/2)-b,$gableht1+$height1-(b*$pitch/12)+5.5)
a = entities.add_line(kathy, kathy+Geom::Vector3d.new($length+$oj1+$oj2,0,0))
a.faces[0].material = $ridge
a = entities.add_line(kathy+Geom::Vector3d.new(0,b*2,0),kathy+Geom::Vector3d.new(0,b*2,0)+Geom::Vector3d.new($length+$oj1+$oj2,0,0))
a.faces[1].material = $ridge


a = entities.add_line([-$oj1+5,($width/2)-b,$gableht1+$height1-(b*$pitch/12)+5.5],[-$oj1+5,-$oh1,$height1+5.5])
a.faces[1].material = $rake
a = entities.add_line(kathy+Geom::Vector3d.new(5,b*2,0),[-$oj1+5,$width+$oh2,$height2+5.5])
a.faces[0].material = $rake
a = entities.add_line(kathy+Geom::Vector3d.new($oj1+$length+$oj2-5,0,0),[$length+$oj2-5,-$oh1,$height1+5.5])
a.faces[0].material = $rake
a = entities.add_line(kathy+Geom::Vector3d.new($oj1+$length+$oj2-5,b*2,0),[$length+$oj2-5,$width+$oh2,$height2+5.5])
a.faces[1].material = $rake

end


entities.add_dimension_linear [0, 0, 0], [$length, 0, 0], [0, $bottom-$width/12/4, 0]
entities.add_dimension_linear [0,0,0],[0,$width,0],[$left-$length*4/12/4,0,0]
a = entities.add_section_plane([0,0,1],[0,0,-1]).activate
a.hidden = true



entities.add_dimension_linear($post_point_EW2[-1], $post_point_SW2[0],$post_point_SW2_vector[0]).arrow_type = Sketchup::Dimension::ARROW_SLASH
entities.add_dimension_linear($post_point_EW1[-1], $post_point_SW1[0],$post_point_SW1_vector[0]).arrow_type = Sketchup::Dimension::ARROW_SLASH
$post_point = $post_point_SW1 + $post_point_EW2 + $post_point_SW2+$post_point_EW1
$post_dimension_direction_vector = $post_point_SW1_vector + $post_point_EW2_vector + $post_point_SW2_vector+$post_point_EW1_vector
$i = 0
while $i<$post_point.length-1 do
	d = entities.add_dimension_linear $post_point[$i], $post_point[$i+1],$post_dimension_direction_vector[$i]
	d.arrow_type = Sketchup::Dimension::ARROW_SLASH
	$i+=1
end

origin = Geom::Point3d.new(-$oj1,-$oh1,-2)
v1 = Geom::Vector3d.new($length+$oj1+$oj2-0.1,0,0)
v2 = Geom::Vector3d.new(0,$width+$oh1+$oh2,0)

entities.add_cline(origin,origin+v1)
entities.add_cline(origin,origin+v2)
entities.add_cline(origin+v1,origin+v1+v2)
entities.add_cline(origin+v2,origin+v2+v1)


def compass(entities)
	pt = Geom::Point3d.new($left-$length*9/12/4,$top,0)
	v1 = Geom::Vector3d.new(20,0,0)
	v2 = Geom::Vector3d.new(0,20,0)
	compass_group = entities.add_group
	
	compass_group.entities.add_line(pt-v2,pt+v2)
	compass_group.entities.add_line(pt+v2+Geom::Vector3d.new(-5,-5,0),pt+v2)
	compass_group.entities.add_line(pt+v2+Geom::Vector3d.new(5,-5,0),pt+v2)
	
	compass_group.entities.add_text "N", pt+v2+Geom::Vector3d.new(0,10,0)
	
	r = Geom::Transformation.rotation pt, [0,0,1], $north.degrees
	entities.transform_entities r, compass_group
	#compass_group.entities.add_text "Salesman: \r\nPhone: \r\nCompany name:", pt-v1-v2
end

# bottom notes



h = $width/4/12
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
t = Geom::Transformation.new [$length*0.7,$bottom-($width/12/2)-h*2,0]
entities.transform_entities t,e
f = entities.add_group
f.entities.add_3d_text('By: '+worksheet.Cells(184,2).Value + '    (888) 466-4139', TextAlignCenter, "Times New Roman", true,false,h,0,-10,true,1)
f.material = "black"
t = Geom::Transformation.new [$length*0.7,$bottom-($width/12/2)-h*3.5,0]
entities.transform_entities t, f
# g = entities.add_group
# g.entities.add_3d_text('PHONE NUMBER: (888) 466-4139', TextAlignCenter, "Times New Roman", true,false,h,0,-10,true,1)
# g.material = "black"
# t = Geom::Transformation.new [$length*0.8,$bottom-($width/12/2)-h*2,0]

# i = entities.add_group
# i.entities.add_3d_text('DATE: ' +worksheet.Cells(183,2).Value.strftime("%m/%d/%Y"), TextAlignCenter, "Times New Roman", true,false,h,0,-10,true,1)
# i.material = "black"
# t = Geom::Transformation.new [$length*0.8,$bottom-($width/12/2)-h*3.5,0]
# entities.transform_entities t, i


#t = Geom::Transformation.new [$left, $bottom, ]
Sketchup.send_action "viewTop:"
Sketchup.active_model.active_view.camera.perspective = false


if($north) 
	compass(entities)
end

rescue NoMemoryError

ensure
	application.Workbooks.Close
end