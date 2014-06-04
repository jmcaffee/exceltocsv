##############################################################################
# File:: exceltocsv.rb
# Purpose:: Include file for ExcelToCsv library
# 
# Author::    Jeff McAffee 04/17/2013
# Copyright:: Copyright (c) 2013, kTech Systems LLC. All rights reserved.
# Website::   http://ktechsystems.com
##############################################################################

require 'find'
require 'logger'


$LOG = Logger.new(STDERR)
$LOG.level = Logger::ERROR

if ENV["DEBUG"] == '1'
  puts "LOGGING: ON due to DEBUG=1"
  $LOG.level = Logger::DEBUG
end

require "#{File.join( File.dirname(__FILE__), 'exceltocsv','version')}"

$LOG.info "**********************************************************************"
$LOG.info "Logging started for ExcelToCsv library."
$LOG.info "**********************************************************************"


require "#{File.join( File.dirname(__FILE__), 'exceltocsv','excel_file')}"

