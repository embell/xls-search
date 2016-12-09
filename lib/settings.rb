def read_setting(setting_to_read)
  settings = read_file('./settings.txt')
  settings.reject! { |line| line.nil? || line.strip.empty? || line.strip[0] == '#' }
  
  settings.each do |line|
    setting = line.split('=').first
    value = line.split('=').last
    return value.strip if setting.strip.upcase == setting_to_read.strip.upcase
  end
  
  nil
end

def read_search_terms
  search_terms = read_file('./search_terms.txt')
  search_terms.reject! { |line| line.nil? || line.strip.empty? || line.strip[0] == '#' }
  
  search_terms
end

def output_folder
  d = Dir.pwd + '/results'
  Dir.mkdir d unless Dir.exists?(d)
  d
end