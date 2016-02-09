     class Asmaa < Sketchup::ModelObserver
       def onPlaceComponent(instance)
            UI.messagebox instance
       end
     end

     Sketchup.active_model.add_observer(Asmaa.new)