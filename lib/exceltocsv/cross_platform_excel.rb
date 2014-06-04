##############################################################################
# File::    cross_platform_excel.rb
# Purpose:: Cross platform Excel binary implementation
#
# Author::    Jeff McAffee 06/03/2014
# Copyright:: Copyright (c) 2014, kTech Systems LLC. All rights reserved.
# Website::   http://ktechsystems.com
##############################################################################
require_relative 'excel_app_wrapper'
require 'spreadsheet'

module ExcelToCsv
  class CrossPlatformExcel < ExcelAppWrapper
    def open_workbook(filepath)
      #   Open an Excel file
      @wb = Spreadsheet.open filepath
    end

    def worksheet_names
      worksheets = @wb.worksheets.collect { |w| w.name }
    end

    def close_workbook
      # NOP
    end

    def worksheet_data(worksheet_name)
      sheet = @wb.worksheet worksheet_name
      sheet.rows
    end
  end
end
