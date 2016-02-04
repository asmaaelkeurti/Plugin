# prompts = ["Wainscot Height (inch)"]
# values = [36]
# popups = [nil]
# results = inputbox(prompts,values,popups)

#results[0]

$wainscot_line = []

$SW1 = Array.new
$SW2 = Array.new
$EW1 = Array.new
$EW2 = Array.new


$SW11 = Array.new
$SW21 = Array.new
$EW11 = Array.new
$EW21 = Array.new




$door_position.each{|x| 	
	$SW1.push(x) if x[0].valid? and (!x[0].instances[0].nil?) and x[0].instances[0].transformation.origin[1] == 0
	$EW2.push(x) if x[0].valid? and (!x[0].instances[0].nil?) and x[0].instances[0].transformation.origin[0] == $length
	$SW2.push(x) if x[0].valid? and (!x[0].instances[0].nil?) and x[0].instances[0].transformation.origin[1] == $width
	$EW1.push(x) if x[0].valid? and (!x[0].instances[0].nil?) and x[0].instances[0].transformation.origin[0] == 0
}

$SW1.sort!{|x,y| x[0].instances[0].transformation.origin[0] <=> y[0].instances[0].transformation.origin[0]}
$EW2.sort!{|x,y| x[0].instances[0].transformation.origin[1] <=> y[0].instances[0].transformation.origin[1]}
$SW2.sort!{|x,y| y[0].instances[0].transformation.origin[0] <=> x[0].instances[0].transformation.origin[0]}
$EW1.sort!{|x,y| y[0].instances[0].transformation.origin[1] <=> x[0].instances[0].transformation.origin[1]}



h = Geom::Vector3d.new(0,0,$wcht)
$SW11.push([0,0,$wcht],[5.5,0,$wcht])
$SW1.each{|x| 	$SW11.push(x[0].instances[0].transformation.origin + h)
				$SW11.push(x[0].instances[0].transformation.origin + Geom::Vector3d.new(0,0,x[2]))
				$SW11.push(x[0].instances[0].transformation.origin + Geom::Vector3d.new(x[1],0,0) + Geom::Vector3d.new(0,0,x[2]))
				$SW11.push(x[0].instances[0].transformation.origin + Geom::Vector3d.new(x[1],0,0) + h)
}
$SW11.push([$length-5.5,0,$wcht],[$length,0,$wcht])


$EW21.push([$length,0,$wcht],[$length,5.5,$wcht])
$EW2.each{|x| 	$EW21.push(x[0].instances[0].transformation.origin + h)
				$EW21.push(x[0].instances[0].transformation.origin + Geom::Vector3d.new(0,0,x[2]))
				$EW21.push(x[0].instances[0].transformation.origin + Geom::Vector3d.new(0,x[1],0) + Geom::Vector3d.new(0,0,x[2]))
				$EW21.push(x[0].instances[0].transformation.origin + Geom::Vector3d.new(0,x[1],0) + h)
}
$EW21.push([$length,$width-5.5,$wcht],[$length,$width,$wcht])

$SW21.push([$length,$width,$wcht],[$length-5.5,$width,$wcht])
$SW2.each{|x| 	$SW21.push(x[0].instances[0].transformation.origin + h)
				$SW21.push(x[0].instances[0].transformation.origin + Geom::Vector3d.new(0,0,x[2]))
				$SW21.push(x[0].instances[0].transformation.origin - Geom::Vector3d.new(x[1],0,0) + Geom::Vector3d.new(0,0,x[2]))
				$SW21.push(x[0].instances[0].transformation.origin - Geom::Vector3d.new(x[1],0,0) + h)
}
$SW21.push([5.5,$width,$wcht],[0,$width,$wcht])

$EW11.push([0,$width,$wcht],[0,$width-5.5,$wcht])
$EW1.each{|x| 	$EW11.push(x[0].instances[0].transformation.origin + h)
				$EW11.push(x[0].instances[0].transformation.origin + Geom::Vector3d.new(0,0,x[2]))
				$EW11.push(x[0].instances[0].transformation.origin - Geom::Vector3d.new(0,x[1],0) + Geom::Vector3d.new(0,0,x[2]))
				$EW11.push(x[0].instances[0].transformation.origin - Geom::Vector3d.new(0,x[1],0) + h)
}
$EW11.push([0,5.5,$wcht],[0,0,$wcht])



entities = Sketchup.active_model.entities

i = 0 
while i+1 < $SW11.length  do
	$wainscot_line.push(entities.add_line($SW11[i],$SW11[i+1]))
	i = i + 1
end

i = 0 
while i+1 < $EW21.length  do
	$wainscot_line.push(entities.add_line($EW21[i],$EW21[i+1]))
	i = i + 1
end

i = 0 
while i+1 < $SW21.length  do
	$wainscot_line.push(entities.add_line($SW21[i],$SW21[i+1]))
	i = i + 1
end

i = 0 
while i+1 < $EW11.length  do
	$wainscot_line.push(entities.add_line($EW11[i],$EW11[i+1]))
	i = i + 1
end