# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)

require 'olm'

Gem::Specification.new do |spec|
  spec.name          = "olm"
  spec.version       = Olm::VERSION
  spec.authors       = ["Takahiro Noda"]
  spec.email         = ["takahiro.noda+rubygems@gmail.com"]
  spec.summary       = %q{A glue library to manipulate MS Outlook Mail.}
  spec.description   = %q{Olm is designed for reading and writing Outlook mail using the win32ole library and against Outlook 2007.}
  spec.homepage      = "https://github.com/tnoda/olm"
  spec.license       = "MIT"
  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.6"
  spec.add_development_dependency "rake"
end
