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

class Overhead < Parametric

# Create windows as components rather than groups
def class_of_object
    Sketchup::ComponentInstance
end

def create_entities(data, container)
    width = data["w"]*12
    height = data["h"]*12
	window= data["n"]
    set = data["s"]
    if $colorOptions.include?(data["c"])
        color = "Kynar " + data["c"] + " Trim"
    else
        color = data["c"] + " Trim"
    end
    Overhead.create_overhead(width, height, window, container, color, set)
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

@@defaults = {"w"=>10, "h"=>10, "n"=>0, "c"=>"Brite","s"=>"set1"}

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
    $colorOptions = ["Brite", "Roman", "Clay", "Beige", "Bronze", "Ash", "Sand", "Autumn", "Tudor", "Smoke", "Evergreen", "BrandyWin", "Terratone", "Matte Black", "Antique Ivory", "Hartford Green"]
    csColorOptions = ["Black","Charcoal","Taupe","Gray","Alamo","Brilliant","Arctic","Forest","Hunter","Gold","Crimson","Rustic","Burgundy","Gallery","Ocean","Ivory","Light Stone","Tan","Brown","Burnished Slate","Copper Metallic"]
    sets = ["set1","set2","set3"]
    prompts = ["Width (feet)*", "Height (feet)*", "Number of Window*", "Color*","set"]
    values = [data["w"], data["h"], data["n"], data["c"], data["s"]]
    popups = [nil, nil,nil,($colorOptions + csColorOptions).join("|"),sets.join("|")]
    results = inputbox( prompts, values, popups, title )
    return nil if not results
    
    # Store the results back into data
    data["w"] = results[0]
    data["h"] = results[1]
	data["n"] = results[2]
    data["c"] = results[3]
    data["s"] = results[4]
    
    # update the defaults values
    if( not @entity )
       data.each {|k, v| @@defaults[k] = v }
    end

    data
end

#-----------------------------------------------------------------------
# Create a rectangular face at a given location


#-----------------------------------
# Create a simple rectangluar frame


#-----------------------------------
# Create a basic window
def Overhead.create_overhead(width, height, window, container, color,set)
    @over_width = width
    @over_height = height
    @set = set

    $irene = container.add_cpoint [0,0,0]
	a = container.add_face [0,0,0],[width,0,0],[width,height,0],[0,height,0]
	a.back_material = color
	a.pushpull 9,true
	a.erase!
	space = height/((height/24).to_i)
	h = space
	while h<height do
		container.add_line([0,h,-9],[width,h,-9])
		h = h + space
	end 
	
	window_width = 24
	window_height = 12
	if(window>0)
			space1 = (width-window*window_width)/(window+1)
			h = space*2 + ((space-12)/2)
			 
			o = Geom::Point3d.new(space1,h,-9)
			v1 = Geom::Vector3d.new(0,window_height,0)
			v2 = Geom::Vector3d.new(window_width,0,0)
			v3 = Geom::Vector3d.new(space1,0,0)
			for i in 1..window
				glass = container.add_face(o,o+v1,o+v1+v2,o+v2)
				glass.material = "[Translucent_Glass_Blue]"
				glass.back_material = "[Translucent_Glass_Blue]"
				o = o+v2+v3
			end
	end
end


#-----------------------------------------------------------------------
# Prompt for parameters and then insert windows
def Overhead.create
    overhead = Overhead.new
    definition= overhead.entity
    Sketchup.active_model.place_component definition
    $door_position.push([definition, @over_width, @over_height, @set,"overhead_"])


end
     

# add a menu with the window types
Overhead.create
#-----------------------------------------------------------------------
end # module Window
