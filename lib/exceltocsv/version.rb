##############################################################################
# Everything is contained in Module ExcelToCsv
#
module ExcelToCsv

  VERSION = "0.1.0" unless constants.include?("VERSION")
  APPNAME = "ExcelToCsv" unless constants.include?("APPNAME")
  COPYRIGHT = "Copyright (c) 2013, kTech Systems LLC. All rights reserved." unless constants.include?("COPYRIGHT")


  def self.logo()
    return  [ "#{ExcelToCsv::APPNAME} v#{ExcelToCsv::VERSION}",
              "#{ExcelToCsv::COPYRIGHT}",
              ""
            ].join("\n")
  end


end # module ExcelToCsv
