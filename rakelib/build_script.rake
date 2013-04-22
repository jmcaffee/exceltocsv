##############################################################################
# File::    build_script.rb
# Purpose:: Build script rake tasks
# 
# Author::    Jeff McAffee 04/15/2013
# Copyright:: Copyright (c) 2013, kTech Systems LLC. All rights reserved.
# Website::   http://ktechsystems.com
##############################################################################

# Clean up the generated script.
# NOTE: These were added to the CLOBBER element and not the CLEAN element
# because the generated script calls 'clean' which would delete itself.
CLOBBER.include("buildgem.")
CLOBBER.include("buildgem.cmd")

# These 'aliases' are added directly to the root namespace
desc "Generate a simple script to build and install this gem"
task :build_gem_script => ["build_script:build_gem_script"]



# Namespaced tasks
namespace :build_script do

  desc "Documentation for building gem"
  task :help do
    hr = "-"*79
    puts hr
    puts "Building the Gem"
    puts "================"
    puts
    puts "Use the following command line to build and install the gem"
    puts
    puts "on *nix"
    puts "rake clean gem && gem install pkg/#{PROJNAME.downcase}-#{PKG_VERSION}.gem -l --no-ri --no-rdoc"
    puts
    puts "on *doze"
    puts "rake clean gem && gem install pkg\\#{PROJNAME.downcase}-#{PKG_VERSION}.gem -l --no-ri --no-rdoc"
    puts
    puts "See also the build_gem_script task which will create a script to build and install the gem."
    puts
    puts hr
  end # task :help


  #desc "Generate a simple script to build and install this gem"
  task :build_gem_script do
    # For windows
    scriptname = "buildgem.cmd"
    if(File.exists?(scriptname))
      puts "Removing existing script."
      rm scriptname
    end

    File.open(scriptname, 'w') do |f|
      f << "::\n"
      f << ":: #{scriptname}\n"
      f << "::\n"
      f << ":: Running this script will generate and install the #{PROJNAME} gem.\n"
      f << ":: Run 'rake build_gem_script' to regenerate this script.\n"
      f << "::\n"

      f << "rake clean gem && gem install pkg\\#{PROJNAME.downcase}-#{PKG_VERSION}.gem -l --no-ri --no-rdoc\n"
    end

    # For nix
    scriptname = "buildgem"
    if(File.exists?(scriptname))
      puts "Removing existing script."
      rm scriptname
    end

    File.open(scriptname, 'w') do |f|
      f << "#\n"
      f << "# #{scriptname}\n"
      f << "#\n"
      f << "# Running this script will generate and install the #{PROJNAME} gem.\n"
      f << "# Run 'rake build_gem_script' to regenerate this script.\n"
      f << "#\n"

      f << "rake clean gem && gem install pkg/#{PROJNAME.downcase}-#{PKG_VERSION}.gem -l --no-ri --no-rdoc\n"
    end

  end # task :build_gem_script
end # namespace :build_script



