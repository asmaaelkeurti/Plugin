     toolbar = UI::Toolbar.new "Test"
     # This toolbar icon simply displays Hello World on the screen
     cmd = UI::Command.new("load excel") {
       load "a/beta.rb"
     }
     cmd.set_validation_proc {
    fileObj = File.new("C:\\Users\\"+ENV['USERNAME']+"\\AppData\\Roaming\\SketchUp\\SketchUp 2015\\SketchUp\\Plugins\\a\\asmaa","r")
    $load = fileObj.gets.include?("1")
    fileObj.close
  if $load
    load 'a/beta.rb'
    fileObj = File.new("C:\\Users\\"+ENV['USERNAME']+"\\AppData\\Roaming\\SketchUp\\SketchUp 2015\\SketchUp\\Plugins\\a\\asmaa","w")
    fileObj.write("0")
    fileObj.close
     MF_GRAYED
   else
     MF_ENABLED
   end
 }
     cmd.small_icon = "ToolPencilSmall.png"
     cmd.large_icon = "ToolPencilLarge.png"
     cmd.tooltip = "Load Excel"
     cmd.status_bar_text = "Testing the toolbars class"
     cmd.menu_text = "Load Excel"
     toolbar = toolbar.add_item cmd
     toolbar.show