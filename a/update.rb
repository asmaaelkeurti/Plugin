require 'win32ole'
require 'sketchup.rb'

begin

	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\test.rb",Sketchup.find_support_file('Plugins'))
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\win32ole.so",Sketchup.find_support_file('Plugins'))
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\fileutils.rb",Sketchup.find_support_file('Plugins'))
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\auto.rb",Sketchup.find_support_file('Plugins'))

	FileUtils.mkdir(Sketchup.find_support_file('Plugins')+'/samples_library') if not File.exist?(Sketchup.find_support_file('Plugins')+'/samples_library')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\samples_library\\parametric.rb",Sketchup.find_support_file('Plugins')+'/samples_library')

	FileUtils.mkdir(Sketchup.find_support_file('Plugins')+'/a') if not File.exist?(Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\asmaa",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\awning.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\basic1.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\basic2.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\beta.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\BLANK3D.skp",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\cross section.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\cupola.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\dormers.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\export.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\floor.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\hydraulic.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\overhead.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\sidelight1.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\sidelight2.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\slideDoor.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\steel length EW1.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\steel length EW2.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\walkdoor.rb",Sketchup.find_support_file('Plugins')+'/a')
	FileUtils.cp("\\\\sv1\\OPEN USE DOCUMENTs\\QUOTEBUILDER\\Sharing Center\\3d drawing\\2016\\a\\windows.rb",Sketchup.find_support_file('Plugins')+'/a')

rescue
	UI.messagebox "Update failed! Please update manually"
else
	UI.messagebox "Update successfully"
end
