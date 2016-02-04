# Copyright 2014 Trimble Navigation Ltd.

require 'sketchup.rb'
require 'extensions.rb'
require 'langhandler.rb'

module Trimble
  module TrimbleConnect

    # Put translation object where the extension can find it. We assign it to a
    # constant so it's accessible from child modules and classes.
    # noinspection RubyConstantNamingConvention
    LH = LanguageHandler.new("trimble_connect.strings")

    # Load the extension.
    extension_name = LH["Trimble Connect"]
    unless Sketchup.is_pro?
      extension_name += " #{LH[" (Pro Only)"]}"
    end

    path = File.dirname(__FILE__).freeze
    loader = File.join(path, "su_trimble_connect", "loader.rb")
    extension = SketchupExtension.new(extension_name, loader)

    extension.description = LH[
        "The Trimble Connect extension allows you to reference, save and "\
        "collaborate on models directly from Trimble Connect. "\
        "Your use of the Trimble Connect extension for SketchUp is subject to "\
        "the Trimble Extension End User License Agreement which can be found "\
        "at: extensions.sketchup.com/en/trimble-extension-eula"]
    extension.version = "1.1.0"
    extension.creator = "SketchUp"
    extension.copyright = "2014-2015, Trimble Navigation Limited"

    # Register the extension with Sketchup.
    Sketchup.register_extension(extension, true)

  end # module TrimbleConnect
end # module Trimble
