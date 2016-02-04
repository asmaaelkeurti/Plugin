# Copyright 2015 Trimble Navigation Ltd.

# Loads the files necessary for operation of the Trimble Connect.

require 'sketchup.rb'


module Trimble::TrimbleConnect

  # Shortcut accessor to make it easy for sub-modules and classes to access this
  # namespace.
  PLUGIN = self

  # Cache the location of the support folder for the extension. This is because
  # getting the path when scrambled can mess things up.
  PATH = File.dirname(__FILE__).freeze

  if Sketchup.is_pro?
    # Load the scrambled extension.
    Sketchup::require "su_trimble_connect/main"
  end

end # module Trimble::TrimbleConnect

