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
# Description :   A tool to create parametric Double Hung and Slider slidedoors.
# Menu Item   :   Tools->Windows
# Context Menu:   Edit Window
# Usage       :   Select slidedoor type and then specify the size.
#             :   If the size needs to be changed after inserting into the model, 
#             :   right click on the slidedoor and select "Edit Window".
# Date        :   9/10/2004
# Type        :   Dialog Box
#-----------------------------------------------------------------------------

# Classes for creating and editing parametric slidedoors

require 'sketchup.rb'
require 'samples_library/parametric.rb'

#=============================================================================
# Define the main parametric slidedoor class

class Slidedoor < Parametric

# Create slidedoors as components rather than groups
def class_of_object
    Sketchup::ComponentInstance
end

def create_entities(data, container)
    width = data["w"]*12
    height = data["h"]*12
    type = data["t"]
    set = data["s"]
    Slidedoor.create_slidedoor(width, height, type, container, set)
end

def create_entity(model)
    #TODO: try to find existing definition matching the parameters
    @entity = model.definitions.add self.compute_name
    
    # Set the behavior
    behavior = @entity.behavior
    behavior.is2d = true
    behavior.snapto = 0
    behavior.cuts_opening = false
    @entity
end

@@defaults = {"w"=>12, "h"=>12, "t"=>0, "s"=>"set1"}

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
    prompts = ["Width (feet)*", "Height (feet)*", "Type*","Set"]
    types = ["Split", "Single left", "Single right"]
    sets = ["set1","set2"]
    values = [data["w"], data["h"], types[data["t"]],data["s"]]
    popups = [nil, nil, types.join("|"),sets.join("|")]
    results = inputbox( prompts, values, popups, title )
    return nil if not results
    
    # Store the results back into data
    data["w"] = results[0]
    data["h"] = results[1]
    t = types.index(results[2])
    data["t"] = t ? t : 0
    data["s"] =  results[3]
    
    # update the defaults values
    if( not @entity )
       data.each {|k, v| @@defaults[k] = v }
    end

    data
end

#-----------------------------------------------------------------------
# Create a rectangular face at a given location

def Slidedoor.rectangle(origin, width, height, container, close)

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
def Slidedoor.simple_frame(p1, width, height, thickness, frameWidth, container)

    # create a group for the frame
    frame = container.add_group
    
    v = Geom::Vector3d.new(frameWidth, frameWidth, 0)
    p2 = p1 + v

    holeWidth = width - (2.0 * frameWidth)
    holeHeight = height - (2.0 * frameWidth)

    # Create the outer frame and the hole
    outer = Slidedoor.rectangle(p1, width, height, frame.entities, true)
    hole = Slidedoor.rectangle(p2, holeWidth, holeHeight, frame.entities, true)
    hole.erase!

    # Extrude the slidedoor
    outer.pushpull(-thickness)

    frame
end

#-----------------------------------
# Create a basic slidedoor
def Slidedoor.create_slidedoor(width, height, type,container, set)
    @door_width = width
    @set = set
#  define same variables
 posX = 0
 posY = 0
 posZ = 0
thickness = 4
tracklen = (2*width)-12
wallColor = "white"
frameColor = "white"
track_color = "white"
wainscot = 0

case type
when 0
   #1 split door
   trackstart = (((width/2)-6) *-1)
when 1
   #2  single   track to the left.  
   trackstart = ((width)*(-1)+12)
else
   #3  single   track to the right
   trackstart = 0
end

#----- comment lines if you dont want a group
   group = container.add_group
   entities = group.entities



     pts = []
     pts[0] = [posX, posY, posZ]
     pts[1] = [(posX+width), posY, posZ]
     pts[2] = [(posX+width), posY, (posZ+height)]
     pts[3] = [posX, posY, (posZ+height)]
   base = entities.add_face pts
   base.material = wallColor
    thickness = -thickness if( base.normal.dot(Z_AXIS) < 0 )
# Now we can do the pushpull
    base.pushpull thickness

	entities.add_line([0,-thickness,height],[0,0,height]).faces[1].material = frameColor
	entities.add_line([0,-thickness,height],[0,0,height]).faces[0].material = frameColor
	entities.add_line([width,-thickness,height],[width,0,height]).faces[0].material = frameColor
	
	p1 = Geom::Point3d.new(0,-thickness,0)
	p2 = Geom::Point3d.new(0,-thickness,height)
	p3 = Geom::Point3d.new(width,-thickness,height)
	p4 = Geom::Point3d.new(width,-thickness,0)
	v1 = Geom::Vector3d.new(-2,0,0)
	v2 = Geom::Vector3d.new(2,0,0)
	v3 = Geom::Vector3d.new(0,0,-2)
	slide_frame = entities.add_face(p1,p2,p3,p4,p4+v1,p3+v1+v3,p2+v2+v3,p1+v2).material = frameColor

 pts = []
 pts[0] = [posX+trackstart, posY, (posZ+height+1)]
 pts[1] = [posX+trackstart, (posY-6), (posZ+height+1)]
 pts[2] = [posX+trackstart, (posY-6), (posZ+height+3)]
 pts[3] = [posX+trackstart, (posY-2), (posZ+height+6)]
 pts[4] = [posX+trackstart, posY, (posZ+height+6)]

 # Add the face to the entities in the model
 face = entities.add_face pts
 face.back_material  = track_color
#  pull track
 face.pushpull -tracklen




#  draw vert lines if a split door
if(wainscot>0.1)
#   draw lines and add them to group
	entities.add_line([2,-$slide_thickness,$slide_wainscot],[$slide_width-2,-$slide_thickness,$slide_wainscot]).faces[1].material = $slide_wainscotcolor
end



case type
when 0

#   draw lines and add them to group
 point1 = Geom::Point3d.new((posX+(width/2)), posY, posZ)
 point2 = Geom::Point3d.new((posX+(width/2)), (posY-thickness), posZ)
 line = entities.add_line point1,point2

 point1 = Geom::Point3d.new((posX+(width/2)), (posY-thickness), posZ)
 point2 = Geom::Point3d.new((posX+(width/2)), (posY-thickness), (posZ+height))
 line = entities.add_line point1,point2

 point1 = Geom::Point3d.new((posX+(width/2)), (posY-thickness), (posZ+height))
 point2 = Geom::Point3d.new((posX+(width/2)), posY, (posZ+height))
 line = entities.add_line point1,point2
end

#  draw wc lines

	r = Geom::Transformation.rotation [0,0,0], [1,0,0],270.degrees
	entities.transform_entities r, group
	#t = Geom::Transformation.new [50,0,0]
	#entities.transform_entities t, group

	group.explode
	
end

#-----------------------------------------------------------------------
# Prompt for parameters and then insert slidedoors
def Slidedoor.create
    slidedoor = Slidedoor.new
    definition= slidedoor.entity
    Sketchup.active_model.place_component definition

    $slideDoor_position.push([definition, @door_width,@set])
end

# add a menu with the slidedoor types
Slidedoor.create
#-----------------------------------------------------------------------
end # module Window
