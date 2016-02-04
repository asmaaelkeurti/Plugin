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

class Window1 < Parametric

# Create windows as components rather than groups
def class_of_object
    Sketchup::ComponentInstance
end

def create_entities(data, container)
	colorOptions = ["Brite", "Roman", "Clay", "Beige", "Bronze", "Ash", "Sand", "Autumn", "Tudor", "Smoke", "Evergreen", "BrandyWin", "Terratone", "Matte Black", "Antique Ivory", "Hartford Green"]
    width = data["w"]*12
    height = data["h"]*12
    type = data["t"]
	grid = data["g"]
	above = data["a"]
	if colorOptions.include?(data["c"])
		color = "Kynar " + data["c"] + " Trim"
	else
		color = data["c"] + " Trim"
	end
    Window1.create_window(width, height, type, grid, container, color, above)
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

@@defaults = {"w" => 4, "h" => 3, "t" => 1,"g"=>"No","c"=>"Brite","a"=>44}

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
    prompts = ["Width (Feet)", "Height (Feet)", "Type* ","Grid* ","Color* ","Height Above Ground"]
    types = ["Verticle", "Slider", "Fixed"]
	boolean = ["Yes","No"]

	csColorOptions = ["Brite", "Roman", "Clay", "Beige", "Bronze", "Ash", "Sand", "Autumn", "Tudor", "Smoke", "Evergreen", "BrandyWin", "Terratone", "Matte Black", "Antique Ivory", "Hartford Green","Black","Charcoal","Taupe","Gray","Alamo","Brilliant","Arctic","Forest","Hunter","Gold","Crimson","Rustic","Burgundy","Gallery","Ocean","Ivory","Light Stone","Tan","Brown","Burnished Slate","Copper Metallic"]
    values = [data["w"], data["h"], types[data["t"]],data["g"],data["c"],data["a"]]
    popups = [nil, nil, types.join("|"),boolean.join("|"), csColorOptions.join("|"),nil]
    results = inputbox( prompts, values, popups, title )
    return nil if not results
    
    # Store the results back into data
    data["w"] = results[0]
    data["h"] = results[1]
    t = types.index(results[2])
    data["t"] = t ? t : 0
	data["g"] = results[3]
	data["c"] = results[4]
	data["a"] = results[5]
    
    # update the defaults values
    if( not @entity )
       data.each {|k, v| @@defaults[k] = v }
    end

    data
end

#-----------------------------------------------------------------------
# Create a rectangular face at a given location

def Window1.rectangle(origin, width, height, container, close)



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
def Window1.simple_frame(p1, width, height, thickness, frameWidth, container)

    # create a group for the frame
    frame = container.add_group
    
    v = Geom::Vector3d.new(frameWidth, frameWidth, 0)
    p2 = p1 + v

    holeWidth = width - (2.0 * frameWidth)
    holeHeight = height - (2.0 * frameWidth)

    # Create the outer frame and the hole
    outer = Window1.rectangle(p1, width, height, frame.entities, true)
    hole = Window1.rectangle(p2, holeWidth, holeHeight, frame.entities, true)
    hole.erase!

    # Extrude the window
    outer.pushpull(-thickness)

    frame
end

#-----------------------------------
# Create a basic window
def Window1.create_window(width, height, type, grid,container, color, above)
		@above = above
		@width = width
		@height = height
		@type = type
		@grid = grid
		@color = color



	w = 1.5
	grid = (grid == "Yes")
	
    origin = Geom::Point3d.new(0,0,0)
	v1 = Geom::Vector3d.new(0,height,0)
	v2 = Geom::Vector3d.new(width,0,0)
	p1 = origin + v1
	p2 = origin + v1 + v2
	p3 = origin + v2
	a = container.add_face origin, p1, p2, p3
	a.back_material = color
	a.material = color

	origin = origin + Geom::Vector3d.new(w,w,0)
	v1 = Geom::Vector3d.new(0,height-2*w,0)
	v2 = Geom::Vector3d.new(width-2*w,0,0)  
	p1 = origin + v1
	p2 = origin + v1 + v2
	p3 = origin + v2
	b = container.add_face origin, p1, p2, p3
	b.material = "[Translucent_Glass_Blue]"
	b.back_material = "[Translucent_Glass_Blue]"

	a.pushpull -w
	
	if(type == 0)
		p = origin
		
		origin = origin + Geom::Vector3d.new(0.1,0,0.1)
		v1 = Geom::Vector3d.new(0,(height-2*w)/2,0)
		v2 = Geom::Vector3d.new(width-2*w,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		c = container.add_face origin, p1, p2, p3
		c.back_material = color

		origin = origin + Geom::Vector3d.new(w,w,0)
		v1 = Geom::Vector3d.new(0,((height-2*w)/2) - 2*w,0)
		v2 = Geom::Vector3d.new(width-4*w,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		d = container.add_face origin, p1, p2, p3
		d.back_material = "[Translucent_Glass_Blue]"
		d.material = "[Translucent_Glass_Blue]"
		
		c.pushpull -1.2
		
		
		origin = p + Geom::Vector3d.new(0,(height-2*w)/2,1.4)
		v1 = Geom::Vector3d.new(0,(height-2*w)/2,0)
		v2 = Geom::Vector3d.new(width-2*w,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		e = container.add_face origin, p1, p2, p3
		e.back_material = color

		origin = origin + Geom::Vector3d.new(w,w,0)
		v1 = Geom::Vector3d.new(0,((height-2*w)/2) - 2*w,0)
		v2 = Geom::Vector3d.new(width-2*w-2*w,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		f = container.add_face origin, p1, p2, p3
		f.back_material = "[Translucent_Glass_Blue]"
		f.material = "[Translucent_Glass_Blue]"
		
		e.pushpull -1.2
	elsif(type == 1)	
		p = origin
		
		origin = origin + Geom::Vector3d.new(0,0,-0.1)
		v1 = Geom::Vector3d.new(0,height-2*w,0)
		v2 = Geom::Vector3d.new((width-2*w)/2,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		c = container.add_face origin, p1, p2, p3
		c.back_material = color

		origin = origin + Geom::Vector3d.new(w,w,0)
		v1 = Geom::Vector3d.new(0,height-4*w,0)
		v2 = Geom::Vector3d.new((width-2*w)/2-w*2,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		d = container.add_face origin, p1, p2, p3
		d.back_material = "[Translucent_Glass_Blue]"
		d.material = "[Translucent_Glass_Blue]"
		
		c.pushpull -1.2
		
		origin = p + Geom::Vector3d.new((width-2*w)/2,0,1.4)
		v1 = Geom::Vector3d.new(0,height-2*w,0)
		v2 = Geom::Vector3d.new((width-2*w)/2,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		e = container.add_face origin, p1, p2, p3
		e.back_material = color
		
		
		
		origin = origin + Geom::Vector3d.new(w,w,0)
		v1 = Geom::Vector3d.new(0,height-4*w,0)
		v2 = Geom::Vector3d.new((width-2*w)/2-w*2,0,0)  
		p1 = origin + v1
		p2 = origin + v1 + v2
		p3 = origin + v2
		f = container.add_face origin, p1, p2, p3
		f.back_material = "[Translucent_Glass_Blue]"
		f.material = "[Translucent_Glass_Blue]"
		
		e.pushpull -1.2
	end
	
	if grid
		origin = Geom::Point3d.new(0,0,0) + Geom::Vector3d.new(w,w,-0.1)
		container.add_line(origin + Geom::Vector3d.new((width -2*w)/4,0,0), origin + Geom::Vector3d.new((width -2*w)/4,0,0) + Geom::Vector3d.new(0,height-2*w,0))
		container.add_line(origin + Geom::Vector3d.new((width -2*w)*3/4,0,0), origin + Geom::Vector3d.new((width -2*w)*3/4,0,0) + Geom::Vector3d.new(0,height-2*w,0))
		container.add_line(origin + Geom::Vector3d.new(0,(height-2*w)/4,0), origin + Geom::Vector3d.new(0,(height-2*w)/4,0) +  Geom::Vector3d.new(width -2*w,0,0))
		container.add_line(origin + Geom::Vector3d.new(0,(height-2*w)*3/4,0), origin + Geom::Vector3d.new(0,(height-2*w)*3/4,0) +  Geom::Vector3d.new(width -w*2,0,0))
	end
	
end

#-----------------------------------------------------------------------
# Prompt for parameters and then insert windows
def Window1.create
    window = Window1.new
    definition= window.entity

    entities = Sketchup.active_model.entities
    $c_line.push(entities.add_cline([0,0,@above],[$length,0,@above]))
    $c_line.push(entities.add_cline([$length,0,@above],[$length,$width,@above]))
    $c_line.push(entities.add_cline([$length,$width,@above],[0,$width,@above]))
    $c_line.push(entities.add_cline([0,$width,@above],[0,0,@above]))


    $window_data.push([definition,@above,@height,@width,@type,@color,@grid,"drag"])

    Sketchup.active_model.place_component definition


end
     


# add a menu with the window types
Window1.create
#-----------------------------------------------------------------------
end # module Window
