# Copyright 2004-2005, @Last Software, Inc.

# This software is provided as an example of using the Ruby interface
# to SketchUp.

# Permission to use, copy, modify, and distribute this software for 
# any purpose and without fee is hereby granted, provided that the above
# copyright notice appear in all copies.

# THIS SOFTWARE IS PROVIDED "AS IS" AND WITHOUT ANY EXPRESS OR
# IMPLIED WARRANTIES, INCLUDING, WITHOUT LIMITATION, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE.
#-----------------------------------------------------------------------------
# Name        :   Window Maker 1.0
# Description :   A tool to create parametric Double Hung and Slider windows.
# Menu Item   :   Tools->Windows
# Context Menu:   Edit Window
# Usage       :   Select window type and then specify the size.
#             :   If the size needs to be changed after inserting into the model, 
#             :   right click on the window and select "Edit Window".
# Date        :   9/10/2004
# Type        :   Dialog Box
#-----------------------------------------------------------------------------

# Classes for creating and editing parametric windows

require 'sketchup.rb'
require 'samples_library/parametric.rb'

#=============================================================================
# Define the main parametric window class

class Dormer < Parametric



  $wcht = 0
$gambrel_height = 0
$concrete_height = 0
$concrete_color = "grey"
$wallColor = "white"
$faciaColor = "white"
$apron = 0
$heel = 8
$gableht1 = (((0.5 * $width)+$oh1)*$pitch)/12
$gableht2 = (((0.5 * $width)+$oh2)*$pitch)/12
$wcht = $wcht + $concrete_height

$posX = 0
$posY = 0
$posZ = 0



    $facia = 5.5
    $corner = 5.5


  
$height1 = $height
$height2 = $height
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
# Create windows as components rather than groups
def class_of_object
    Sketchup::ComponentInstance
end

def create_entities(data, container)
	colorOptions = ["Brite", "Roman", "Clay", "Beige", "Bronze", "Ash", "Sand", "Autumn", "Tudor", "Smoke", "Evergreen", "BrandyWin", "Terratone", "Matte Black", "Antique Ivory", "Hartford Green"]
    $width = data["w"]*12

    $height = data["h"]*12

    $length = data["l"]*12
	$pitch = data["p"]
	$oh1 = data["oh"]*12
	$oh2 = $oh1
	$oj1 = data["oj"]*12
	$oj2 = $oj1





  
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
    Dormer.create_window(container)
end

def create_entity(model)
    #TODO: try to find existing definition matching the parameters
    @entity = model.definitions.add self.compute_name
    
    # Set the behavior
    behavior = @entity.behavior
    behavior.is2d = true
    behavior.snapto = 0
    behavior.cuts_opening = true
    @entity
end

@@defaults = {"w" => 6, "h" => 6, "l" => 6,"p"=>6,"oh"=>1,"oj"=>1}

def default_parameters
    @@defaults.clone
end

def translate_key(key)
    prompt = key
    case( key )
        when "w"
            prompt = "Width"
        when "h"
            prompt = "Height"
    end
    prompt
end

# Show a dialog and get the values from the user
# The default implementation of this in the Parametric class doesn't support
# having a popup list.  Maybe I should consider adding something that would
# allow doing that in a more generic way.
def prompt(operation)
    # get the parameters
    if( @entity )
        data = self.parameters
    else
        data = self.default_parameters
    end
    if( not data )
        puts "No parameters attached to the entity"
        return nil
    end
    title = "(* values will not import)"
    #title = operation + " " + self.class.name
    prompts = ["Width (Feet)", "Height (Feet)", "length","pitch","Overhang ","Overjet "]



	csColorOptions = ["Brite", "Roman", "Clay", "Beige", "Bronze", "Ash", "Sand", "Autumn", "Tudor", "Smoke", "Evergreen", "BrandyWin", "Terratone", "Matte Black", "Antique Ivory", "Hartford Green","Black","Charcoal","Taupe","Gray","Alamo","Brilliant","Arctic","Forest","Hunter","Gold","Crimson","Rustic","Burgundy","Gallery","Ocean","Ivory","Light Stone","Tan","Brown","Burnished Slate","Copper Metallic"]
    values = [data["w"], data["h"],data["l"],data["p"],data["oh"],data["oj"]]
    popups = [nil, nil, nil, nil,nil,nil]
    results = inputbox( prompts, values, popups, title )
    return nil if not results
    
    # Store the results back into data
    data["w"] = results[0]
    data["h"] = results[1]
    data["l"] = results[2]
    data["p"] = results[3]
	data["oh"] = results[4]
	data["oj"] = results[5]
    
    # update the defaults values
    if( not @entity )
       data.each {|k, v| @@defaults[k] = v }
    end

    data
end

#-----------------------------------------------------------------------
# Create a rectangular face at a given location

def Dormer.rectangle(origin, width, height, container, close)



    v1 = Geom::Vector3d.new(width,0,0)
    v2 = Geom::Vector3d.new(0,height,0)
    p1 = origin;
    p2 = origin + v1
    p3 = p2 + v2
    p4 = origin + v2

    edges = []
    edges[0]=container.add_line p1, p2
    edges[1]=container.add_line p2, p3
    edges[2]=container.add_line p3, p4
    edges[3]=container.add_line p4, p1

    if( close )
        f = container.add_face edges
    else
        edges
    end
    
end

#-----------------------------------
# Create a simple rectangluar frame
def Dormer.simple_frame(p1, width, height, thickness, frameWidth, container)

    # create a group for the frame
    frame = container.add_group
    
    v = Geom::Vector3d.new(frameWidth, frameWidth, 0)
    p2 = p1 + v

    holeWidth = width - (2.0 * frameWidth)
    holeHeight = height - (2.0 * frameWidth)

    # Create the outer frame and the hole
    outer = Dormer.rectangle(p1, width, height, frame.entities, true)
    hole = Dormer.rectangle(p2, holeWidth, holeHeight, frame.entities, true)
    hole.erase!

    # Extrude the window
    outer.pushpull(-thickness)

    frame
end

#-----------------------------------
# Create a basic window
def Dormer.create_window(container)

entities = container
basic_group = entities.add_group
entities = basic_group.entities
 
 
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
  
  
  
  
  # add_3d_letter_ew1 entities, "G", logo_y_value, logo_z_value
  # add_3d_letter_ew2 entities, "B", logo_y_value -7.5, logo_z_value, $length

  # add_3d_letter_ew1 entities, "B", logo_y_value - 7.5, logo_z_value - 7.5
  # add_3d_letter_ew2 entities, "G", logo_y_value - 15, logo_z_value + 7.5, $length



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
    # draw corner lines on rear end wall and horz line at top-Geom
  #   entities.add_line([$posX, ($posY + $corner), $posZ],[$posX, ($posY + $corner), ($posZ + $height1)])
  #   entities.add_line([$posX, ($posY + $corner), ($posZ + $height1)],[$posX, ($posY), ($posZ + $height1)])
  # 
  #   entities.add_line([$posX, ($posY + $width-$corner), $posZ],[$posX, ($posY + $width-$corner), ($posZ + $height2)])
  #   entities.add_line([$posX, ($posY + $width-$corner), ($posZ + $height2)],[$posX, ($posY + $width), ($posZ + $height2)])

    # draw $corner lines on front end wall and horz line at top
    # entities.add_line([($posX + $length), ($posY + $corner), $posZ],[($posX + $length), ($posY + $corner), ($posZ + $height1)])
    # entities.add_line([($posX + $length), ($posY + $corner), ($posZ + $height1)],[($posX + $length), ($posY), ($posZ + $height1)])

    # entities.add_line([($posX + $length), ($posY + $width-$corner), $posZ],[($posX + $length), ($posY + $width-$corner), ($posZ + $height2)])
    # entities.add_line([($posX + $length), ($posY + $width-$corner), ($posZ + $height2)],[($posX + $length), ($posY + $width), ($posZ + $height2)])

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
  # entities.add_line([($posX + $corner), ($posY + $width), $posZ],[($posX + $corner), ($posY + $width), ($posZ + $height2)])
  # entities.add_line([($posX + $length-$corner), ($posY + $width), $posZ],[($posX + $length-$corner), ($posY + $width), ($posZ + $height2)])

  # draw $corner lines on back side wall (y=0)
  # entities.add_line([($posX + $corner), $posY, $posZ],[($posX + $corner), $posY, ($posZ + $height1)])
  # entities.add_line([($posX + $length-$corner), $posY, $posZ],[($posX + $length-$corner), $posY, ($posZ + $height1)])

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

if($wcht>0)		
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
transformation = Geom::Transformation.rotation([0,0,0],[0,0,1],90.degrees)
Sketchup.active_model.entities.transform_entities transformation, basic_group
end

#-----------------------------------------------------------------------
# Prompt for parameters and then insert windows
def Dormer.create
    window = Dormer.new
    definition= window.entity
    definition.behavior.snapto = SnapTo_Arbitrary
    Sketchup.active_model.place_component definition


end
     


# add a menu with the window types
Dormer.create
#-----------------------------------------------------------------------
end # module Window
