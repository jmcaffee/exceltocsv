##############################################################################
# File::    win_excel.rb
# Purpose:: Windows OLE Excel Application Wrapper
#
# Author::    Jeff McAffee 06/03/2014
# Copyright:: Copyright (c) 2014, kTech Systems LLC. All rights reserved.
# Website::   http://ktechsystems.com
##############################################################################
require_relative 'excel_app_wrapper'
require 'win32ole'

module ExcelToCsv
  class WinExcel < ExcelAppWrapper
    def open_workbook(filepath)
      #   Open an Excel file
      @xl = WIN32OLE.new('Excel.Application')
      # Turn off excel alerts.
      @xl.DisplayAlerts = false

      # 2nd param of false turns off the link update request
      # when an xls file is opened that contains links.
      @wb = @xl.Workbooks.Open("#{filepath}", false)
    end

    def worksheet_names
      worksheets = []
      @wb.Worksheets.each do |ws|
        worksheets << ws.Name
      end
      worksheets
    end

    def close_workbook
      # Close Excel
      @xl.Quit
    end

    def worksheet_data(worksheet_name)
      data = @wb.Worksheets(worksheet_name).UsedRange.Value
    end
  end
end
