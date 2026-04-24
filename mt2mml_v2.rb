#!/usr/bin/env ruby
# mt2mml_v2.rb
# Usage: ruby mt2mml_v2.rb /path/to/oleObject.bin
#
# Output: MathML string to stdout (or empty string)

in_path = ARGV[0]
unless in_path && File.exist?(in_path)
  puts ""
  exit 0
end

def safe_require(name)
  require name
  true
rescue LoadError
  false
end

def normalize_mathml(s)
  return "" if s.nil?
  x = s.to_s

  # strip BOM + whitespace
  x = x.sub(/\A\uFEFF/, "").strip

  # remove xml header if any
  x = x.gsub(/<\?xml[^>]*\?>/i, "").strip

  # ensure <math> has xmlns
  if x =~ /<math\b/i && x !~ /<math\b[^>]*\bxmlns=/i
    x = x.sub(/<math\b/i, '<math xmlns="http://www.w3.org/1998/Math/MathML"')
  end

  # menclose radical -> msqrt (cứu căn trong 1 số output)
  # do multiple passes to handle nested
  re = /<menclose\b[^>]*\bnotation\s*=\s*["']radical["'][^>]*>(.*?)<\/menclose>/im
  5.times do
    break unless x =~ re
    x = x.gsub(re, '<msqrt>\1</msqrt>')
  end

  x.strip
end

def try_convert_plus(file_path)
  return nil unless safe_require("mathtype_to_mathml_plus")

  # Tùy gem version: Converter/convert API có thể khác nhau
  if defined?(MathTypeToMathMLPlus::Converter)
    c = MathTypeToMathMLPlus::Converter.new(file_path)
    return c.convert
  end

  # fallback API variants
  if defined?(MathTypeToMathMLPlus) && MathTypeToMathMLPlus.respond_to?(:convert)
    return MathTypeToMathMLPlus.convert(file_path)
  end

  nil
rescue => _e
  nil
end

def try_convert_base(file_path)
  return nil unless safe_require("mathtype_to_mathml")

  if defined?(MathTypeToMathML::Converter)
    c = MathTypeToMathML::Converter.new(file_path)
    return c.convert
  end

  if defined?(MathTypeToMathML) && MathTypeToMathML.respond_to?(:convert)
    return MathTypeToMathML.convert(file_path)
  end

  nil
rescue => _e
  nil
end

mml = nil

# 1) ưu tiên plus
mml = try_convert_plus(in_path)

# 2) fallback base
mml = try_convert_base(in_path) if mml.nil? || mml.to_s.strip.empty?

# 3) normalize output
out = normalize_mathml(mml)

# Ensure we only print a single <math>...</math> block if multiple appear
# (hiếm nhưng có thể xảy ra nếu converter trả thêm wrapper)
if out.include?("<math") && out.include?("</math>")
  # take first math block
  start_idx = out.index("<math")
  end_idx = out.index("</math>", start_idx)
  if start_idx && end_idx
    out = out[start_idx, (end_idx - start_idx + "</math>".length)]
  end
end

puts(out)
exit 0
