require 'roo'
require 'roo-xls'
require 'axlsx'
require 'colorize'
require 'fileutils'
require 'pry'

require_relative './lib/generic.rb'
require_relative './lib/settings.rb'
require_relative './lib/search.rb'
require_relative './lib/report.rb'

at_exit do
  puts "\nPress enter to exit."
  gets.chomp
end

begin
  DATA_DIRECTORY = read_setting('Data Folder')
  read_setting('Ignore Case').upcase == 'TRUE' ? IGNORE_CASE = true : IGNORE_CASE = false
rescue Errno::ENOENT
  puts "Verify that settings.txt file exists.".red
  exit
end

unless Dir.exist?(DATA_DIRECTORY)
  puts "Invalid directory: [#{DATA_DIRECTORY}]. Check your settings file.".red
  exit
end

begin
  search_terms = read_search_terms
rescue Errno::ENOENT
  puts "Search terms file was not found. Verify that search_terms.txt is present in this folder.".red
  exit
end

if search_terms.empty?
  puts "No search terms found. Check your search_terms.txt file.".red
  exit
end

if IGNORE_CASE
  search_terms.map! { |term| term.upcase }
end

results = parse_datasheets(search_terms, DATA_DIRECTORY, IGNORE_CASE)

puts "\nResults".green
puts "-------".green
results.each do |key, value|
  puts "#{key.to_s.cyan} - #{value.count.to_s.cyan} sheets"
end

create_report(results)