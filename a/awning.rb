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

class Awning < Parametric

# Create windows as components rather than groups
def class_of_object
    Sketchup::ComponentInstance
end

def create_entities(data, container)
    width = data["w"]*12
    type = data["t"]
	length = data["l"]*12
	pitch = data["p"]
    set = data["s"]
    Awning.create_awning(width, type, length, pitch, container,set)
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

@@defaults = {"w"=>10, "t"=>1,"l"=>3,"p"=>6.0,"s"=>"Awning#1"}

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
    prompts = ["Length*", "Type*","Width*","pitch*","Set"]
    types = ["Gable", "Hip"]
	boolean = ["Yes","No"]
    sets = ["Awning#1","Awning#2","Awning#3"]
    values = [data["w"], types[data["t"]],data["l"],data["p"],data["s"]]
    popups = [nil, types.join("|"),nil,nil,sets.join("|")]
    results = inputbox( prompts, values, popups, title )
    return nil if not results
    
    # Store the results back into data
    data["w"] = results[0]
    t = types.index(results[1])
    data["t"] = t ? t : 0
	data["l"] = results[2]
	data["p"] = results[3]
    data["s"] = results[4]
    
    # update the defaults values
    if( not @entity )
       data.each {|k, v| @@defaults[k] = v }
    end

    data
end

#-----------------------------------------------------------------------
# Create a rectangular face at a given location

def Awning.rectangle(origin, width, height, container, close)

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
def Awning.simple_frame(p1, width, height, thickness, frameWidth, container)

    # create a group for the frame
    frame = container.add_group
    
    v = Geom::Vector3d.new(frameWidth, frameWidth, 0)
    p2 = p1 + v

    holeWidth = width - (2.0 * frameWidth)
    holeHeight = height - (2.0 * frameWidth)

    # Create the outer frame and the hole
    outer = awning.rectangle(p1, width, height, frame.entities, true)
    hole = awning.rectangle(p2, holeWidth, holeHeight, frame.entities, true)
    hole.erase!

    # Extrude the window
    outer.pushpull(-thickness)

    frame
end

#-----------------------------------
# Create a basic window
def Awning.create_awning(width, type, length, pitch, container, set)
    @width = width
    @set = set



  h = length*pitch/12

  if(type == 1)
	asmaa = length
  else
	asmaa = 0
  end
  
  x1 = 6
  y1 = 6
  x = 6
  y = 6
  c = (8.5-6)/2
	pt1 = Geom::Point3d.new(0,0,0)
	pt2 = Geom::Point3d.new(width, 0, 0)
	pt3 = Geom::Point3d.new(width, length, 0)
	pt4 = Geom::Point3d.new(0, length, 0)
	pt5 = Geom::Point3d.new(asmaa,0,h)
	pt6 = Geom::Point3d.new(width-asmaa,0,h)
	v1 = Geom::Vector3d.new(0,0,-5.5)
	
	porch_group = container.add_group
	porch_group.entities.add_line pt5,pt6
	face = porch_group.entities.add_face(pt1,pt2,pt3,pt4)
	face.pushpull 5.5
	
	#porch_group.container.add_face(pt1+v1,pt2+v1,pt3+v1,pt4+v1).material = $porch_ceiling_color
	
	
	
	porch_group.entities.add_line pt2, pt6
	
	porch_group.entities.add_line pt4, pt5
	porch_group.entities.add_line(pt1, pt5).find_faces
    porch_group.entities.add_line(pt3, pt6).find_faces

=begin
	if(asmaa>0)
		porch_group.entities.add_face(pt1,pt4,pt5).material = $roofColor
		porch_group.entities.add_face(pt2,pt6,pt3).material = $roofColor
	else
		porch_group.entities.add_face(pt1,pt4,pt5).material = $wallColor
		porch_group.entities.add_face(pt2,pt6,pt3).material = $wallColor
	end
	porch_group.entities.add_face(pt6,pt5,pt4,pt3).material = $roofColor
=end

r = Geom::Transformation.rotation [0, 0, 0], [1,0,0], 270.degrees	
container.transform_entities r, porch_group
r = Geom::Transformation.rotation [0, 0, 0], [0,1,0], 180.degrees
container.transform_entities r, porch_group
end

#-----------------------------------------------------------------------
# Prompt for parameters and then insert windows
def Awning.create
    awning = Awning.new
    definition= awning.entity
    Sketchup.active_model.place_component definition

    $awning_data.push([definition,@width,@set])
end

# add a menu with the window types
Awning.create
#-----------------------------------------------------------------------
end # module Window
