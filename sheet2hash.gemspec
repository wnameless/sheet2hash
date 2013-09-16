# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'sheet2hash/version'
require 'date'

Gem::Specification.new do |spec|
  spec.name          = "sheet2hash"
  spec.version       = Sheet2hash::VERSION
  spec.authors       = ["Wei-Ming Wu"]
  spec.email         = ["wnameless@gmail.com"]
  spec.description   = %q{Convert Excel or Spreadsheet to Ruby hash}
  spec.date          = "#{Date.today.to_s}"
  spec.summary       = "sheet2hash-#{Sheet2hash::VERSION}"
  spec.homepage      = "http://github.com/wnameless/sheet2hash"
  spec.license       = "Apache License, Version 2.0"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_dependency "roo"

  spec.add_development_dependency "simplecov"
  spec.add_development_dependency "yard"
  spec.add_development_dependency "bundler", "~> 1.3"
  spec.add_development_dependency "rake"
end
