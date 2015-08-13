# ExcelToCsv::ExcelFile

ExcelFile is a file converter to convert Excel spreadsheets to CSV files.
It is specifically designed for the criteria required to generate properly
formated CSV files for use with [GDLC](https://github.com/jmcaffee/gdlc).

## Usage

Quick example:

    require 'exceltocsv'

    converter = ExcelToCsv::ExcelFile.new
    converter.xl_to_csv( 'path/to/input.xls', 'path/to/output.csv' )

Example rake task that updates (converts xls files) csvs based on last modified
date of each file within a directory structure.

###### plk.rake

    require 'exceltocsv'

    desc "Update CSV files from XLS source"
    task :update do
      plks = FileList['plk/xls/**/*.xls']

      # Pathmap string maps to csv dir with csv target file
      pm = "%{^plk/xls,plk/csv;.xls$,.csv;.xlsx$,.csv}p"
      # Remove any source files when the dest file exists and is newer.
      plks.delete_if do |s|
        # Downcase the path,
        d = s.pathmap( pm ).downcase
        # and snakecase the target filename.
        d = snakecase_filename(d)
        File.exists?(d) && File.stat(s).mtime <= File.stat(d).mtime
      end

      target_csvs = plks.pathmap( pm )

      # I want the target filenames normalized to lower case.
      target_csvs.each { |p| p.downcase! }

      # Create all target dirs
      target_dirs = target_csvs.pathmap("%d")
      target_dirs.uniq!
      mkdir_p target_dirs

      # Convert all newer XL files to CSVs.
      # Note that this method only converts the first sheet in the workbook.
      converter = ExcelToCsv::ExcelFile.new
      plks.each do |x|
        converter.xl_to_csv(x, snakecase_filename(x.pathmap(pm).downcase))
      end

      puts "All target files are up to date" if plks.empty?
    end

    def snakecase_filename(filepath)
      snake_file_path = File.join(filepath.pathmap("%d"), filepath.pathmap("%n").snakecase + filepath.pathmap("%x"))
    end

## License

See [LICENSE](https://github.com/jmcaffee/exceltocsv/blob/master/LICENSE).
Website: [http://ktechsystems.com](http://ktechsystems.com)

