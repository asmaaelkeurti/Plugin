prompts = ["cupola size (feet)"]
values = [3]
popups = [nil]
results = inputbox(prompts,values,popups)


$cub_size = results[0]*12


$gambrel_height = 0
$roofColor = "white"
$wallColor = "white"

$pitch = 3.5


def build_cupola(a, entities, pitch, t, roof, wall)
		
		cupola_group = entities.add_group
		f1 = cupola_group.entities.add_face([a,0,0],[a,a/2, a*$pitch/2/12], [a, a, 0], [a, a, a*$pitch/2/12 + 4 + a], [a,0,a*$pitch/2/12 + 4 + a])
		f1.pushpull -a

		cupola_group.entities.add_edges([0,0,a*$pitch/2/12 + 4],[a,0,a*$pitch /2/12+ 4 ])[0].faces[1].material = roof
		cupola_group.entities.add_edges([0,0,a*$pitch/2/12 + 4],[a,0,a*$pitch/2/12 + 4 ])[0].faces[0].material = wall
		cupola_group.entities.add_edges([a,0,a*$pitch/2/12 + 4],[a,a,a*$pitch/2/12 + 4 ])[0].faces[1].material = roof
		cupola_group.entities.add_edges([a,0,a*$pitch/2/12 + 4],[a,a,a*$pitch/2/12 + 4 ])[0].faces[0].material = wall
		cupola_group.entities.add_edges([a,a,a*$pitch/2/12 + 4],[0,a,a*$pitch/2/12 + 4])[0].faces[1].material = roof
		cupola_group.entities.add_edges([a,a,a*$pitch/2/12 + 4],[0,a,a*$pitch/2/12 + 4])[0].faces[0].material = wall
		cupola_group.entities.add_edges([0,0,a*$pitch/2/12 + 4],[0,a,a*$pitch/2/12 + 4])[0].faces[0].material = roof
		cupola_group.entities.add_edges([0,0,a*$pitch/2/12 + 4],[0,a,a*$pitch/2/12 + 4])[0].faces[1].material = wall
		
		o = 2
		
		face = cupola_group.entities.add_face([-o,-o,a*$pitch/2/12 + 4 + a],[a/2,a/2, a*3/2+ a*8/2/12 + 4], [-o,a+o,a*$pitch/2/12 + 4 + a])
		face.material = roof
		face = cupola_group.entities.add_face([-o,a+o,a*$pitch/2/12 + 4 + a],[a/2,a/2, a*3/2 + a*8/2/12 + 4],[a+o,a+o,a*$pitch/2/12 + 4 + a])
		face.material = roof
		face = cupola_group.entities.add_face([a+o,a+o,a*$pitch/2/12 + 4 + a],[a/2,a/2, a*3/2 + a*8/2/12 + 4],[a+o,0-o,a*$pitch/2/12 + 4 + a])
		face.material = roof
		face =  cupola_group.entities.add_face([a+o,0-o,a*$pitch/2/12 + 4 + a],[a/2,a/2, a*3/2 + a*8/2/12 + 4],[0-o,0-o,a*$pitch/2/12 + 4 + a])
		face.material = roof
		face = cupola_group.entities.add_face([-o,-o,a*$pitch/2/12 + 4 + a],[-o,a+o,a*$pitch/2/12 + 4 + a],[a+o,a+o,a*$pitch/2/12 + 4 + a],[a+o,0-o,a*$pitch/2/12 + 4 + a])
		face.material = roof
		pt = Geom::Point3d.new(a/2,a/2, a*3/2+ a*8/2/12 + 4 + a/2)
		cupola_group.entities.add_line(pt,pt+Geom::Vector3d.new(0,0,-a*3/4))
		cupola_group.entities.add_line(pt,pt+Geom::Vector3d.new(a/4,0,0))
		cupola_group.entities.add_line(pt,pt+Geom::Vector3d.new(-a/4,0,0))
		cupola_group.entities.add_line(pt,pt+Geom::Vector3d.new(0,a/4,0))
		cupola_group.entities.add_line(pt,pt+Geom::Vector3d.new(0,-a/4,0))
		cupola_group.entities.add_line(pt,pt+Geom::Vector3d.new(0,0,a/4))
		entities.transform_entities t, cupola_group


end


entities = Sketchup.active_model.entities




		t = Geom::Transformation.new [-200,-200, 0]
		build_cupola($cub_size, entities, $pitch, t, $roofColor, $wallColor)

  
