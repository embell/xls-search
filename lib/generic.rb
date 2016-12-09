def get_percent_complete(completed, total)
  return "Cannot get percentage when total is zero! total:[#{total}]" if total == 0
  percentage = completed.to_f / total.to_f * 100.to_f
  "#{percentage.round(1)}% complete...\t".rjust(5)
end

def read_file(file)
  return "File is too large to read" if File.size(file) > 5000000
  results = Array.new
  f = File.open(file, 'r')
  f.each { |l| results << l.chomp }
  f.close
  results
end