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
# Description :   A tool to create parametric Double Hung and Slider Hydraulics.
# Menu Item   :   Tools->Windows
# Context Menu:   Edit Window
# Usage       :   Select Hydraulic type and then specify the size.
#             :   If the size needs to be changed after inserting into the model, 
#             :   right click on the Hydraulic and select "Edit Window".
# Date        :   9/10/2004
# Type        :   Dialog Box
#-----------------------------------------------------------------------------

# Classes for creating and editing parametric Hydraulics

require 'sketchup.rb'
require 'samples_library/parametric.rb'

#=============================================================================
# Define the main parametric Hydraulic class

class Hydraulic < Parametric

# Create Hydraulics as components rather than groups
def class_of_object
    Sketchup::ComponentInstance
end

def create_entities(data, container)
    width = data["w"]*12
    height = data["h"]*12
    type = data["t"]
    Hydraulic.create_Hydraulic(width, height,container)
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

@@defaults = {"w"=>12, "h"=>12}

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
    prompts = ["Width (feet)*", "Height (feet)*"]
    values = [data["w"], data["h"]]
    popups = [nil, nil]
    results = inputbox( prompts, values, popups, title )
    return nil if not results
    
    # Store the results back into data
    data["w"] = results[0]
    data["h"] = results[1]
    
    # update the defaults values
    if( not @entity )
       data.each {|k, v| @@defaults[k] = v }
    end

    data
end

#-----------------------------------------------------------------------
# Create a rectangular face at a given location

def Hydraulic.rectangle(origin, width, height, container, close)

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
def Hydraulic.simple_frame(p1, width, height, thickness, frameWidth, container)

    # create a group for the frame
    frame = container.add_group
    
    v = Geom::Vector3d.new(frameWidth, frameWidth, 0)
    p2 = p1 + v

    holeWidth = width - (2.0 * frameWidth)
    holeHeight = height - (2.0 * frameWidth)

    # Create the outer frame and the hole
    outer = Hydraulic.rectangle(p1, width, height, frame.entities, true)
    hole = Hydraulic.rectangle(p2, holeWidth, holeHeight, frame.entities, true)
    hole.erase!

    # Extrude the Hydraulic
    outer.pushpull(-thickness)

    frame
end

#-----------------------------------
# Create a basic Hydraulic
def Hydraulic.create_Hydraulic(width, height,container)
@width = width

wainscot = 0
hydraulic_color = "white"
wainscot_color = "white"
	hydraulic_group = container.add_group
f = hydraulic_group.entities.add_face [0,0,6],[0,-4, 6],[0,-4,6+height],[0,0,8+6+height]
b = hydraulic_group.entities.add_line([0,0,6+height],[0,-4,6+height])
b.faces[0].material = "white"

if (wainscot >0 and wainscot < height)
	w = hydraulic_group.entities.add_line([0,0,6+wainscot],[0,-4,6+wainscot]) 
	b.faces[0].pushpull width, true
	b.faces[0].back_material = "white"
	w.find_faces
	w.faces[0].material = hydraulic_color

	w.faces[1].material = wainscot_color
	w.faces[1].pushpull width, true
	w.faces[1].back_material = wainscot_color
	w.faces[0].pushpull width, true
	w.faces[0].back_material = hydraulic_color
end
if (wainscot == height)
	b.faces[1].material = $hydraulic_wainscot_color
	b.faces[1].pushpull $hydraulic_width, true
	b.faces[1].back_material = $hydraulic_wainscot_color
	b.faces[0].pushpull $hydraulic_width, true
	b.faces[0].back_material = "black"
end
if (wainscot == 0)
	b.faces[1].material = hydraulic_color
	b.faces[1].pushpull width, true
	b.faces[1].back_material = wainscot_color
	b.faces[0].pushpull width, true
	b.faces[0].back_material = "white"
end
r = Geom::Transformation.rotation [0, 0, 0], [1,0,0], 270.degrees
container.transform_entities r, hydraulic_group

hydraulic_group.explode

end

#-----------------------------------------------------------------------
# Prompt for parameters and then insert Hydraulics
def Hydraulic.create
    hydraulic = Hydraulic.new
    definition= hydraulic.entity
    Sketchup.active_model.place_component definition
    $hydraulic_position.push([definition, @width,"set"])
end

# add a menu with the slidedoor types
Hydraulic.create
#-----------------------------------------------------------------------
end # module Window
