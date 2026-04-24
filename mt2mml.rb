# mt2mml.rb - Simple MathType to MathML converter using gem
require 'mathtype_to_mathml'

path = ARGV[0]
abort "usage: ruby mt2mml.rb <oleObject*.bin>" unless path && File.exist?(path)

begin
  converter = MathTypeToMathML::Converter.new(path)
  puts converter.convert
rescue => e
  STDERR.puts "Error: #{e.message}"
  exit 1
end
