class Result
    attr_reader :file, :worksheet
  
  def initialize(file, worksheet)
    @file = file
    @worksheet = worksheet
  end
  
  def ==(other)
    other.class == self.class && other.worksheet.upcase == self.worksheet.upcase && other.file.upcase == self.file.upcase
  end
end

def parse_datasheets(keywords, directory, ignore_case)
  all_spreadsheet_names = Dir.glob("#{File.expand_path(directory)}/*.xls")
  
  puts "Searching in the following directory: ".green
  puts directory.cyan
  puts "which expands to: ".green
  puts File.expand_path(directory).cyan
  puts
  puts "Searching for the following keywords:".green
  puts keywords.inspect.cyan
  puts
  puts "Searching through ".green + all_spreadsheet_names.count.to_s.cyan + " spreadsheets".green
  
  worksheet_count = 0
  errors = []
  results_per_keyword = Hash[ keywords.map { |keyword| [keyword, []] } ]
  
  # Loop over each file
  all_spreadsheet_names.each_with_index do |file_name, file_index|
    spreadsheet = open_spreadsheet(file_name)
    
    # Loop over each worksheet
    spreadsheet.each_with_pagename do |sheet_name, sheet|
      # Loop over each row
      begin
        sheet.each do |row|
          # Loop over each cell in the row
          row.each do |cell_text|
            cell_text = cell_text.to_s
            if ignore_case
              cell_text.upcase!
            end
            unless cell_text.nil? || cell_text.empty?
              if keywords.include?(cell_text)
                add_result_for_keyword(cell_text, file_name, sheet_name, results_per_keyword)
              end
            end
          end
        end
        worksheet_count += 1
      rescue ArgumentError # An exception gets thrown when an empty sheet is encountered
        errors << "Sheet [#{sheet_name}] in file [#{file_name}]"
      end
    end
    
    print "\r#{get_percent_complete(file_index+1, all_spreadsheet_names.count)}".green
  end
  
  if errors.any?
    puts "\n\nThe following sheets could not be searched, most likely because they are empty: "
    errors.each do |e|
      puts "- #{e.red}"
    end
  end
  
  puts "\nExamined ".green + worksheet_count.to_s.cyan + " worksheets.".green
  
  results_per_keyword
end

def open_spreadsheet(file_name)
  # roo-xls uses a deprecated method that causes a warning. Temporarily suppress warnings.
  original_warning_level = $VERBOSE
  $VERBOSE = nil
  
  spreadsheet = Roo::Excel.new(file_name)
  
  # Reset warning level
  $VERBOSE = original_warning_level
  
  spreadsheet
end

def add_result_for_keyword(keyword, file, sheet_name, results_per_keyword)
  current_result = Result.new(file, sheet_name)
  
  results_per_keyword[keyword] << current_result unless results_per_keyword[keyword].include?(current_result)
end