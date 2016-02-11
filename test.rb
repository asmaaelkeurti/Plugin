require 'win32ole'

Sketchup.template = Sketchup.find_support_file('Plugins') + "/a/BLANK3D.skp"

plugins_menu = UI.menu("Plugins")
plugins_menu.add_separator
plugins_menu.add_separator
plugins_menu.add_separator
plugins_menu.add_separator
plugins_menu.add_item("Load Excel") {load 'a/beta.rb'}
plugins_menu.add_item("Export Window Data") {load 'a/export.rb'}
plugins_menu.add_separator
plugins_menu.add_item("Windows"){load "a/windows.rb"}
plugins_menu.add_separator

sub3d = plugins_menu.add_submenu("3D Items")
sub3d.add_item("Overhead"){load "a/overhead.rb"}
sub3d.add_item("Walkdoor"){load "a/walkdoor.rb"}
sub3d.add_item("Awning"){load "a/awning.rb"}
sub3d.add_item("Hydraulic"){load "a/hydraulic.rb"}
sub3d.add_item("SlideDoor"){load "a/slideDoor.rb"}
sub3d.add_item("Dormer"){load "a/dormers.rb"}
sub3d.add_item("Cupola"){load "a/cupola.rb"}
sub3d.add_separator
sub3d.add_item("Basic Building 1"){load "a/basic1.rb"}
sub3d.add_item("Basic Building 2"){load "a/basic2.rb"}

sub3d.add_separator
sub3d.add_item("Sidelight1"){load "a/sidelight1.rb"}
sub3d.add_item("Sidelight2"){load "a/sidelight2.rb"}


sub2d = plugins_menu.add_submenu("2D Drawing")
sub2d.add_item("steel length EW1"){load "a/steel length EW1.rb"}
sub2d.add_item("steel length EW2"){load "a/steel length EW2.rb"}
sub2d.add_item("cross section"){load "a/cross section.rb"}
sub3d.add_item("floor plan"){load "a/floor.rb"}
#plugins_menu.add_item("Wainscot"){load "a/wainscot.rb"}
#plugins_menu.add_item("Wainscot2"){load "a/wainscot2.rb"}



$door_position = []
$wainscot_line = []
$window_data = []
$c_line = []
$slideDoor_position = []
$hydraulic_position = []
$basic_building = []
$awning_data = []
$wcht = 0


class NewModel < Sketchup::AppObserver
   def onOpenModel(model)
     model.add_observer(WindowCline.new)
     model.add_observer(CutWcht.new)
   end
end



class WindowCline < Sketchup::ModelObserver
   def onPlaceComponent(instance)
   		$c_line.each{|x| x.erase! if x.valid?} 
        Sketchup.active_model.remove_observer WindowCline
   end
end


class CutWcht < Sketchup::ModelObserver
    def onPlaceComponent(instance)
        if instance.definition.name["walkdoor"]
            entities = Sketchup.active_model.entities
            height = instance.definition.get_attribute("walkdoor","height",80)
            width = instance.definition.get_attribute("walkdoor","width",36)

            p = instance.transformation.origin

            cutWainscot(p,width,height,entities,instance)

        elsif instance.definition.name["overhead"]
            entities = Sketchup.active_model.entities
            height = instance.definition.get_attribute("overhead","height",120)
            width = instance.definition.get_attribute("overhead","width",120)

            p = instance.transformation.origin

            cutWainscot(p,width,height,entities,instance)
        end
    end
end

def cutWainscot(p,width,height,entities,instance)
    Sketchup.break_edges = true
    if p[0] == 0
        p1 = instance.transformation.origin + Geom::Vector3d.new(0,0,$wcht)
        p2 = Geom::Point3d.new(p1) + Geom::Vector3d.new(0,0,height-$wcht)
         entities.add_line(p1,p2) 

        p3 = Geom::Point3d.new(p2) - Geom::Vector3d.new(0,width,0)
        entities.add_line(p2,p3)

        p4 = Geom::Point3d.new(p3) - Geom::Vector3d.new(0,0,height-$wcht)
        entities.add_line(p3,p4)

        p5 = Geom::Point3d.new(p1) - Geom::Vector3d.new(0,width,0)
        entities.add_line(p1,p5).erase! if $wcht > 0
    elsif p[0] == $length

        p1 = instance.transformation.origin + Geom::Vector3d.new(0,0,$wcht)
        p2 = Geom::Point3d.new(p1) + Geom::Vector3d.new(0,0,height-$wcht)
        entities.add_line(p1,p2)

        p3 = Geom::Point3d.new(p2) + Geom::Vector3d.new(0,width,0)
        entities.add_line(p2,p3)

        p4 = Geom::Point3d.new(p3) - Geom::Vector3d.new(0,0,height-$wcht)
        entities.add_line(p3,p4)

        p5 = Geom::Point3d.new(p1) + Geom::Vector3d.new(0,width,0)
        entities.add_line(p1,p5).erase! if $wcht > 0


    elsif p[1] == 0

        p1 = instance.transformation.origin + Geom::Vector3d.new(0,0,$wcht)
        p2 = Geom::Point3d.new(p1) + Geom::Vector3d.new(0,0,height-$wcht)
        entities.add_line(p1,p2) 
        p3 = Geom::Point3d.new(p2) + Geom::Vector3d.new(width,0,0)
        entities.add_line(p2,p3)

        p4 = Geom::Point3d.new(p3) - Geom::Vector3d.new(0,0,height-$wcht)
        entities.add_line(p3,p4)

        p5 = Geom::Point3d.new(p1) + Geom::Vector3d.new(width,0,0)
        entities.add_line(p1,p5).erase! if $wcht > 0


    elsif p[1] == $width

        p1 = instance.transformation.origin + Geom::Vector3d.new(0,0,$wcht)
        p2 = Geom::Point3d.new(p1) + Geom::Vector3d.new(0,0,height-$wcht)
        entities.add_line(p1,p2) if $wcht > 0

        p3 = Geom::Point3d.new(p2) - Geom::Vector3d.new(width,0,0)
        entities.add_line(p2,p3)

        p4 = Geom::Point3d.new(p3) - Geom::Vector3d.new(0,0,height-$wcht)
        entities.add_line(p3,p4)

        p5 = Geom::Point3d.new(p1) - Geom::Vector3d.new(width,0,0)
        entities.add_line(p1,p5).erase! if $wcht > 0

    end
end


Sketchup.add_observer(NewModel.new)
Sketchup.active_model.add_observer(CutWcht.new)    
Sketchup.active_model.add_observer(WindowCline.new)






      toolbar = UI::Toolbar.new "test"
     # This toolbar icon simply displays Hello World on the screen
     cmd = UI::Command.new("Export") {
       output
     }
     cmd.set_validation_proc {
    fileObj = File.new("C:\\Program Files (x86)\\Google\\Google SketchUp 8\\Plugins\\a\\asmaa","r")
    $load = fileObj.gets.include?("2")
    fileObj.close
  if $load
    fileObj = File.new("C:\\Program Files (x86)\\Google\\Google SketchUp 8\\Plugins\\a\\asmaa","w")
    fileObj.write("0")
    fileObj.close
    output

     MF_GRAYED
   else
     MF_ENABLED
   end
 }
     cmd.small_icon = "ToolPencilSmall.png"
     cmd.large_icon = "ToolPencilLarge.png"
     cmd.tooltip = "Export model"
     cmd.status_bar_text = "Testing the toolbars class"
     cmd.menu_text = "Export"
     toolbar = toolbar.add_item cmd
     toolbar.show


def output
#window
                application = WIN32OLE.new('Excel.Application')
            workbook = application.Workbooks.Open("C:\\Program Files (x86)\\Google\\Google SketchUp 8\\Plugins\\a\\test")
            worksheet=workbook.Worksheets("Sheet2")
            
            worksheet.Range("A200:Z250").Clear

            i = 1
            $window_data.each{|x|
                if x[0].valid? and (!x[0].instances[0].nil?)
                    i = i + 1
                    
                    a = x[0].instances[0].transformation.origin
                    worksheet.Cells(202,i).Value = a[2].to_i
                    worksheet.Cells(203,i).Value = x[2]
                    worksheet.Cells(204,i).Value = x[3]

                    if x[7] == "spreedsheet"
                        worksheet.Cells(205,i).Value = x[4]
                        if x[6] = 0
                            worksheet.Cells(207,i).Value = "No"
                        else
                            worksheet.Cells(207,i).Value = "Yes"
                        end


                    else
                        if x[4] == 0
                            worksheet.Cells(205,i).Value = "Verticle"
                        elsif x[4] == 1
                            worksheet.Cells(205,i).Value = "Slider"
                        else
                            worksheet.Cells(205,i).Value = "Fixed"
                        end
                        worksheet.Cells(207,i).Value = x[6]
                    end


                    
                    worksheet.Cells(206,i).Value = x[5]
                    
                    if a[0] == 0
                        worksheet.Cells(200,i).Value = "EW1"
                        worksheet.Cells(201,i).Value = $width - a[1].to_i
                    elsif a[0] == $length
                        worksheet.Cells(200,i).Value = "EW2"
                        if x[7] == "spreedsheet"
                            worksheet.Cells(201,i).Value = a[1].to_i - x[3]
                        else
                            worksheet.Cells(201,i).Value = a[1].to_i
                        end
                    elsif  a[1] == 0
                        worksheet.Cells(200,i).Value = "SW1"
                        worksheet.Cells(201,i).Value = a[0].to_i
                    else
                        worksheet.Cells(200,i).Value = "SW2"        
                        worksheet.Cells(201,i).Value = $length - a[0].to_i
                    end


                end

            }
            worksheet.Cells(208,1).Value = i - 1


#walkdoor and overhead

            i = 1
            $door_position.each{|x|
                if x[0].valid? and (!x[0].instances[0].nil?)
                    a = x[0].instances[0].transformation.origin
                    i = i + 1
                    if a[0] == 0
                        worksheet.Cells(210,i).Value = "Endwall 1"
                        worksheet.Cells(211,i).Value = $width - a[1].to_i
                    elsif a[0] == $length
                        worksheet.Cells(210,i).Value = "Endwall 2"
                        worksheet.Cells(211,i).Value = a[1].to_i
                    elsif  a[1] == 0
                        worksheet.Cells(210,i).Value = "Sidewall 1"
                        worksheet.Cells(211,i).Value = a[0].to_i
                    else
                        worksheet.Cells(210,i).Value = "Sidewall 2"        
                        worksheet.Cells(211,i).Value = $length - a[0].to_i
                    end
                    worksheet.Cells(212,i).Value = x[4] + x[3]
                end
            }


#slideDoor
            i = 1
            $slideDoor_position.each{|x|
                if x[0].valid? and (!x[0].instances[0].nil?)
                    a = x[0].instances[0].transformation.origin
                    i = i + 1
                    if a[0] == 0
                        worksheet.Cells(214,i).Value = "EW1"
                        worksheet.Cells(215,i).Value = $width - a[1].to_i
                    elsif a[0] == $length
                        worksheet.Cells(214,i).Value = "EW2"
                        worksheet.Cells(215,i).Value = a[1].to_i
                    elsif  a[1] == 0
                        worksheet.Cells(214,i).Value = "SW1"
                        worksheet.Cells(215,i).Value = a[0].to_i
                    else
                        worksheet.Cells(214,i).Value = "SW2"        
                        worksheet.Cells(215,i).Value = $length - a[0].to_i
                    end
                    worksheet.Cells(216,i).Value = x[2]
                end
            }

            i = 1
            $hydraulic_position.each{|x|
                if x[0].valid? and (!x[0].instances[0].nil?)
                    a = x[0].instances[0].transformation.origin
                    i = i + 1
                    if a[0] == 0
                        worksheet.Cells(218,i).Value = "EW1"
                        worksheet.Cells(219,i).Value = $width - a[1].to_i - x[1]
                    elsif a[0] == $length
                        worksheet.Cells(218,i).Value = "EW2"
                        worksheet.Cells(219,i).Value = a[1].to_i - x[1]
                    elsif  a[1] == 0
                        worksheet.Cells(218,i).Value = "SW1"
                        worksheet.Cells(219,i).Value = a[0].to_i-x[1]
                    else
                        worksheet.Cells(218,i).Value = "SW2"        
                        worksheet.Cells(219,i).Value = $length - a[0].to_i - x[1]
                    end 
                    worksheet.Cells(220,i).Value = x[2]

                end
            }

            if $basic_building.length > 0
                worksheet.Cells(222,2).Value = $basic_building[0]
                worksheet.Cells(223,2).Value = $basic_building[1]
                worksheet.Cells(224,2).Value = $basic_building[2]
                worksheet.Cells(225,2).Value = $basic_building[3]
                worksheet.Cells(226,2).Value = $basic_building[4]
                worksheet.Cells(227,2).Value = $basic_building[5]
                worksheet.Cells(228,2).Value = $basic_building[6]
                worksheet.Cells(229,2).Value = $basic_building[7]
                worksheet.Cells(230,2).Value = $basic_building[8]

            end    

            i = 1
            $awning_data.each{|x|
                if x[0].valid? and (!x[0].instances[0].nil?)
                    a = x[0].instances[0].transformation.origin
                    i = i + 1
                    if a[0] == 0
                        worksheet.Cells(232,i).Value = "EW1"
                        worksheet.Cells(233,i).Value = $width - a[1].to_i - x[1]
                        worksheet.Cells(235,i).Value = x[2]+"EW"
                    elsif a[0] == $length
                        worksheet.Cells(232,i).Value = "EW2"
                        worksheet.Cells(233,i).Value = a[1].to_i - x[1]
                        worksheet.Cells(235,i).Value = x[2]+"EW"
                    elsif  a[1] == 0
                        worksheet.Cells(232,i).Value = "SW1"
                        worksheet.Cells(233,i).Value = a[0].to_i-x[1]
                        worksheet.Cells(235,i).Value = x[2]+"SW"
                    else
                        worksheet.Cells(232,i).Value = "SW2"        
                        worksheet.Cells(233,i).Value = $length - a[0].to_i - x[1]
                        worksheet.Cells(235,i).Value = x[2]+"SW"
                    end 
                    worksheet.Cells(234,i).Value = a[2]-5.5
                    

                end
            }

            # if !$sidelight1.nil?
            #     worksheet.Cells(237,2).Value = $sidelight1
            



                workbook.save
                application.Workbooks.Close
                application.quit
end


























