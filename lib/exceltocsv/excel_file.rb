##############################################################################
# File::    excel_file.rb
# Purpose:: Convert Excel files to CSV format accounting for formating
#           This file was originally located in the gdlrakeutils gem
#           as converttocsv.rb
#
# Author::    Jeff McAffee 04/17/2013
# Copyright:: Copyright (c) 2013, kTech Systems LLC. All rights reserved.
# Website::   http://ktechsystems.com
##############################################################################

require_relative 'os'
require 'time'
require 'csv'
require 'bigdecimal'


module ExcelToCsv
class ExcelFile

  # If FORCE_WIN_OLE is Y or 1, use the actual Excel application.
  # NOTE: This will only work on a windows OS!
  unless ENV['FORCE_WIN_OLE'].nil?
    FORCE_WIN_OLE = 1 if ENV['FORCE_WIN_OLE'] == '1'
    FORCE_WIN_OLE = 1 if ENV['FORCE_WIN_OLE'].downcase == 'y'
  end

  def initialize()
    @date_RE = Regexp.new(/\d{4,4}\/\d{2,2}\/\d{2,2}/)
    @date_with_dashes_RE = Regexp.new(/\d{4,4}-\d{2,2}-\d{2,2}/)
    @date_with_time_RE = Regexp.new(/\d{2,2}:\d{2,2}:\d{2,2}/)
  end

  def xl_app
    return @xl_app unless @xl_app.nil?
    if OS.windows? and defined?(FORCE_WIN_OLE)
      require_relative 'win_excel'
      @xl_app = WinExcel.new
    else
      # CrossPlatformExcel is faster (like, by 30x).
      require_relative 'cross_platform_excel'
      @xl_app = CrossPlatformExcel.new
    end
    @xl_app
  end

  def set_flag(flg)
    if (flg == "-v")
      @verbose = true
    end
  end

  def verbose?()
    @verbose ||= false
  end

  # Convert the 1st sheet in an xls(x) file to a csv file.
  def xl_to_csv(infile, outfile)
    filepath = File.expand_path(infile)
    puts "xl_to_csv: #{infile} => #{outfile}" if verbose?

    unless File.exists?(filepath)
      puts "Unable to find file."
      puts "  #{filepath}"
      return
    end

    #   Open an Excel file
    xl_app.open_workbook filepath

    # Build a list of work sheets to dump to file.
    sheets_in_file = []

    sheet_saved_count = 0

    xl_app.worksheet_names.each do |sheetname|
      if( sheetname.match(/CQDS/) || sheetname.match(/PLK/) )
        sheets_in_file << sheetname
        puts "Converting sheet #{sheetname}" if verbose?
        sheet_saved_count += 1
      end
    end

    if (1 > sheet_saved_count)
      puts "*** No sheets labeled 'PLK' or 'CQDS' ***"
      puts "Verify #{infile} is formatted correctly."
      # Close Excel
      xl_app.close_workbook
      return
    end

    # Write sheet data to file.
    File.open(outfile, "w") do |f|
      data = xl_app.worksheet_data(sheets_in_file[0])
      for row in data
        row_data = []
        for a_cell in row
          row_data << process_cell_value(a_cell)
        end

        contains_data = false

        # Determine if the row contains any data.
        for cell in row_data
          if(cell.match(/[^,\r\n]+/))
            contains_data = true
          end
        end

        # Insert an empty line if the row contains no data.
        if(true == contains_data)
          f << row_data.join(",")
          f << "\n"

          if(true == verbose?)
            puts "#{row_data}"
          end

        else
          f << "\n"

          if(true == verbose?)
            puts "\n"
          end
        end
      end
    end

    # Strip empty data from end of lines
    clean_csv(outfile)

    # Close Excel
    xl_app.close_workbook
  end


  def clean_csv(filename)
    max_row_length = 0
    CSV.foreach(filename) do |row|
      row_len = 0
      i = 0
      row.each do |item|
        row_len = i if !item.nil? && !item.empty?
        i += 1
      end
      max_row_length = row_len if row_len > max_row_length
    end

    puts "Max row length: #{max_row_length.to_s}" if verbose?

    tmp_file = filename.to_s + ".tmp.csv"
    CSV.open(tmp_file, "wb") do |tmp_csv|
      # Used to track empty lines
      empty_found = false

      CSV.foreach(filename) do |row|
        i = 0
        clean_row = []
        while(i <= max_row_length) do
          clean_row << row[i]
          i += 1
        end
        # We need to stop output on 2nd empty row
        break if empty_row?(clean_row) && empty_found
        empty_found = empty_row?(clean_row)
        tmp_csv << clean_row
      end # CSV read
    end # CSV write

    # Replace original file with tmpfile.
    FileUtils.rm filename
    FileUtils.mv tmp_file, filename
  end

  # Return true if row contains no data
  def empty_row?(row)
    is_empty = true
    row.each do |item|
      is_empty = false if item && !item.empty?
    end
    is_empty
  end

  def process_cell_value(a_cell)
    # Truncate the number to 3 decimal places if numeric.
    a_cell = truncate_decimal(a_cell)

    # Remove leading and trailing spaces.
    a_cell = a_cell.to_s.strip

    # If the result is n.000... Remove the unecessary zeros.
    a_cell = clean_int_value(a_cell)

    # If the result is a date, remove time.
    a_cell = format_date(a_cell)

    # Surround the cell value with quotes when it contains a comma.
    a_cell = '"' + a_cell + '"' if a_cell.include?(',')

    a_cell
  end

  # Truncates a decimal to 3 decimal places if numeric
  # and remove trailing zeros, if more than one decimal place.
  # returns a string
  def truncate_decimal(a_cell)
    if(a_cell.is_a?(Numeric))
      a_cell = truncate_decimal_to_string(a_cell, 3)
      # Truncate zeros (unless there is only 1 decimal place)
      # eg. 12.10 => 12.1
      #     12.0  => 12.0
      a_cell = BigDecimal.new(a_cell).to_s("F")
    end
    a_cell
  end

  # Truncates a decimal and converts it to a string.
  # num: decimal to truncate
  # places: number of decimal places to truncate at
  def truncate_decimal_to_string(num, places)
    "%.#{places}f" % num
  end

  # If the result is n.000... Remove the unecessary zeros.
  def clean_int_value(a_cell)
    if(a_cell.match(/\.[0]+$/))
      cary = a_cell.split(".")
      a_cell = cary[0]
    end
    a_cell
  end

  # If the cell is a date, format it to MM/DD/YYYY, stripping time.
  def format_date(a_cell)
    isdate = true if(nil != (dt = a_cell.match(@date_RE)))
    isdate = true if(isdate || (nil != (dt = a_cell.match(@date_with_dashes_RE))) )
    isdate = true if(isdate || (nil != (dt = a_cell.match(@date_with_time_RE))) )
    if isdate
      begin
        mod_dt = DateTime.parse(a_cell)
        cary = "#{mod_dt.month}/#{mod_dt.day}/#{mod_dt.year}"
        if(true == verbose?)
          puts ""
          puts "*** Converted date to #{cary} ***"
          puts ""
        end
        a_cell = cary
      rescue ArgumentError => e
        # Either this is not a date, or the date format is unrecognized,
        # nothing to see here, moving on.
      end
    end
    a_cell
  end

  def prepare_outdir(outdir)
    if( !File.directory?(outdir) )
      FileUtils.makedirs("#{outdir}")
    end
  end

  def winPath(filepath)
    parts = filepath.split("/")
    mspath = nil

    for part in parts
      if(mspath == nil)
        mspath = []
        mspath << part
      else
        mspath << "\\" << part
      end
    end

    mspath
  end
end # class ExcelFile
end # module ExcelToCsv
