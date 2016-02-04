entities = Sketchup.active_model.entities
$sidelight_down = 24

origin = Geom::Point3d.new($corner, 0, $height1)
v1 = Geom::Vector3d.new($length-2*$corner,0,0)
v2 = Geom::Vector3d.new(0,0,$sidelight_down)
pt1 = origin + v1
pt2 = origin + v1 - v2
pt3 = origin - v2
a = entities.add_face(origin,pt1,pt2,pt3)
a.erase!
b = entities.add_face origin,pt1,pt2,pt3
b.material = "[Translucent_Glass_Blue]"
b.back_material = "[Translucent_Glass_Blue]"