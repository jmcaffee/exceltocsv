##############################################################################
# File::    excel_app_wrapper.rb
# Purpose:: Excel Application Wrapper base class
#
# Author::    Jeff McAffee 06/03/2014
# Copyright:: Copyright (c) 2014, kTech Systems LLC. All rights reserved.
# Website::   http://ktechsystems.com
##############################################################################

module ExcelToCsv
  class ExcelAppWrapper
    def open_workbook(filepath)
      fail 'abstract #open_workbook method must be overridden'
    end

    def worksheet_names
      fail 'abstract #worksheet_names method must be overridden'
    end

    def close_workbook
      fail 'abstract #close_workbook method must be overridden'
    end

    def worksheet_data(worksheet_name)
      fail 'abstract #worksheet_data method must be overridden'
    end
  end
end
