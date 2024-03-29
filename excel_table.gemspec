# coding: utf-8
lib = File.expand_path("../lib", __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require "excel_table/version"

Gem::Specification.new do |spec|
  spec.name = "excel_table"
  spec.version = ExcelTable::VERSION
  spec.authors = ["Kenneth Leung"]
  spec.email = ["kenneth@leungs.us"]
  spec.description = %q{A simple gem to generate an Excel report}
  spec.summary = %q{A simple gem to generate an Excel report}
  spec.homepage = ""
  spec.license = "MIT"

  spec.files = `git ls-files`.split($/)
  spec.executables = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler"
  spec.add_development_dependency "rake"
  spec.add_development_dependency "rspec"
  spec.add_runtime_dependency "caxlsx"
end
