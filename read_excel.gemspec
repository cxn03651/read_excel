require_relative 'lib/read_excel/version'

Gem::Specification.new do |spec|
  spec.name          = "read_excel"
  spec.version       = ReadExcel::VERSION
  spec.authors       = ["Hideo NAKAMURA"]
  spec.email         = ["nakamura.hideo@gmail.com"]

  spec.summary       = %q{utility for reading excel file both .xls and .xlsx.}
  spec.description   = %q{utility for reading excel file both .xls and .xlsx.}
  spec.license       = "MIT"
  spec.homepage      = "https://github.com/cxn03651/read_excel"
  spec.required_ruby_version = Gem::Requirement.new(">= 2.3.0")

  # Specify which files should be added to the gem when it is released.
  # The `git ls-files -z` loads the files in the RubyGem that have been added into git.
  spec.files         = Dir.chdir(File.expand_path('..', __FILE__)) do
    `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  end
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]
  spec.add_dependency "rubyXL"
  spec.add_dependency "spreadsheet"
end
