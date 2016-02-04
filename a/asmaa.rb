     class WindowCline < Sketchup::ModelObserver
       def onPlaceComponent(instance)
            $c_line.each{|x| x.erase! if x.valid?} 
            Sketchup.active_model.remove_observer WindowCline
       end
     end

 
     class MyModelObserver < Sketchup::ModelObserver
       def onPlaceComponent(instance)
        if $wcht > 1
            load "a/wainscot2.rb"
            load "a/wainscot.rb"
         end

             Sketchup.active_model.remove_observer MyModelObserver
       end
     end
     Sketchup.active_model.add_observer(WindowCline.new)
     Sketchup.active_model.add_observer(MyModelObserver.new)