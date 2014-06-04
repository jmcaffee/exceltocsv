# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'exceltocsv/version'
require 'exceltocsv/os'

Gem::Specification.new do |spec|
  spec.name          = "exceltocsv"
  spec.version       = ExcelToCsv::VERSION
  spec.authors       = ["Jeff McAffee"]
  spec.email         = ["jeff@ktechsystems.com"]
  spec.description   = %q{ExcelToCsv is a utility library for converting Excel files to CSV format.}
  spec.summary       = %q{Utility for converting Excel files to CSV format}
  spec.homepage      = ""
  spec.license       = "Mine"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.3"
  spec.add_development_dependency "rake"
  if OS.windows?
    #spec.add_runtime_dependency "win32ole"
  end
  spec.add_runtime_dependency "spreadsheet"
end
