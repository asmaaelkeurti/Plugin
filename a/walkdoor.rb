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

class Walkdoor < Parametric

# Create windows as components rather than groups
def class_of_object
    Sketchup::ComponentInstance
end

def create_entities(data, container)
    window = data["w"]
    grid = data["g"]
	swing = data["s"]
	set = data["t"]
	if $colorOptions.include?(data["c"])
		color = "Kynar " + data["c"] + " Trim"
	else
		color = data["c"] + " Trim"
	end
    Walkdoor.create_walkdoor(window, grid, swing, container, color, set)
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

@@defaults = {"w"=>"Yes", "g"=>"No", "s"=>"Right","c"=>"Brite","t"=>"set1"}

def default_parameters
    @@defaults.clone
end

def translate_key(key)
    prompt = key
    case( key )
        when "w"
            prompt = "Window"
        when "g"
            prompt = "grid"
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
    prompts = ["Window*", "Grid*", "Knob Location*","Color*","set"]
	boolean = ["Yes","No"]
	direction = ["Right","Left"]
	sets = ["set1","set2","set3"]
	$colorOptions = ["Brite", "Roman", "Clay", "Beige", "Bronze", "Ash", "Sand", "Autumn", "Tudor", "Smoke", "Evergreen", "BrandyWin", "Terratone", "Matte Black", "Antique Ivory", "Hartford Green"]
	csColorOptions = ["Black","Charcoal","Taupe","Gray","Alamo","Brilliant","Arctic","Forest","Hunter","Gold","Crimson","Rustic","Burgundy","Gallery","Ocean","Ivory","Light Stone","Tan","Brown","Burnished Slate","Copper Metallic"]
    values = [data["w"], data["g"], data["s"], data["c"], data["t"]]
    popups = [boolean.join("|"),boolean.join("|"),direction.join("|"), ($colorOptions + csColorOptions).join("|"),sets.join("|")]
    results = inputbox( prompts, values, popups, title )
    return nil if not results
    
    # Store the results back into data
    data["w"] = results[0]
    data["g"] = results[1]
	data["s"] = results[2]
	data["c"] = results[3]
	data["t"] = results[4]
    
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
def Walkdoor.create_walkdoor(window, grid, swing, container,color,set)



	window_height = 37
	reset = 3
	window_width = 22
	window_height = 36
	window_offset = 7
	knob_height = 36
	
	width = 36
	height = 80

	@door_width = 36
	@door_height = 80
	@door_set = set



	a = container.add_face([0,0,0],[width,0,0],[width,height,0],[0,height,0])
	a.back_material = color
	a.pushpull reset, true
	a.erase!
	
	
	o = Geom::Point3d.new(window_offset,window_height,-reset)
	v1 = Geom::Vector3d.new(0,window_height,0)
	v2 = Geom::Vector3d.new(window_width,0,0)
	
	if(window == "Yes")
		b = container.add_face(o,o+v1,o+v1+v2,o+v2)
		b.material = "[Translucent_Glass_Blue]"
		b.back_material = "[Translucent_Glass_Blue]"
	end
	
	if(swing == "Right")
		center_point = Geom::Point3d.new(width-3.5,knob_height,-reset+2)
	else
		center_point = Geom::Point3d.new(3.5,knob_height,-reset+2)
	end
		radius = 2
		normal_vector = Geom::Vector3d.new(0,0,1)
		edgearray = container.add_circle center_point, normal_vector, radius
		edgearray[0].find_faces
		face = edgearray[0].faces[0]
		
		face.pushpull -4
	
	if(grid == "Yes")
		v3 = Geom::Vector3d.new(window_width/3,0,0)
		v4 = Geom::Vector3d.new(0,window_height/3,0)
		container.add_line(o+v3,o+v3+v1)
		container.add_line(o+v3+v3,o+v3+v3+v1)
		container.add_line(o+v4,o+v4+v2)
		container.add_line(o+v4+v4,o+v4+v4+v2)
	end
	
	
	
	
end



#-----------------------------------------------------------------------
# Prompt for parameters and then insert windows
def Walkdoor.create
    walkdoor = Walkdoor.new
    definition= walkdoor.entity
    Sketchup.active_model.place_component definition, false 
    $door_position.push([definition, @door_width, @door_height,@door_set,"walkdoor_"])
    # Sketchup.active_model.add_observer(MyModelObserver.new)
end
     


     # class MyModelObserver < Sketchup::ModelObserver
     #   def onPlaceComponent(instance)
     #   	if $wcht > 0
     #     	load "a/wainscot2.rb"
     #     	load "a/wainscot.rb"
     #     end

     #         Sketchup.active_model.remove_observer MyModelObserver
     #   end
     # end
# add a menu with the window types
Walkdoor.create
#-----------------------------------------------------------------------
end # module Window
