def create_report(all_results)
  
  puts "\nData processing complete, creating Excel file...".green
  
  report_container = Axlsx::Package.new
  report = report_container.workbook
  
  @styles = {
              title: report.styles.add_style(:sz => 16, :b => true, :alignment => {:horizontal => :center}),
              subtitle: report.styles.add_style(:b => true),
              separator: report.styles.add_style(:bg_color => "FFCCCCCC"),
              warning: report.styles.add_style(:fg_color => "FFFF0000"),
              link: report.styles.add_style(:fg_color => "FF0000FF", :u => true)
            }

  all_results.each_with_index do |(user, tests), i|
    add_worksheet(report, user, tests)
    print "\r#{get_percent_complete(i+1, all_results.count)}".green
  end
  puts "\n\nExcel file built! Saving report...".green
  
  filename = File.expand_path("#{output_folder}/xls-search_report_#{Time.now.strftime("%Y-%b-%d_%H-%M-%S")}.xlsx")
  
  report_container.serialize(filename)
  
  puts "\nExcel file creation complete! Opening report.".green
  system("start #{filename}")
  
rescue => e
  puts e.inspect.red
end

def add_worksheet(workbook, keyword, results)
  worksheet_name = get_sheet_name(keyword, workbook)
  
  sheet = workbook.add_worksheet(name: worksheet_name)
  sheet.merge_cells "A1:B1"
  sheet.add_row [keyword, ''], :style=>@styles[:title]
  sheet.add_row ['', ''], :style=>@styles[:separator], :widths => [50, 50]
  
  if results.nil? || results.empty?
    sheet.add_row(["No results found for #{keyword}"])
  else
    sheet.add_row ["Found in #{results.count} sheets:"], :style=>[@styles[:subtitle]]
    results.each do |result|
      sheet.add_row [result.file.split('/').last, result.worksheet]
    end
  end
  
  worksheet_name
end

def get_sheet_name(starting_title, workbook)
  sheet_name = starting_title[0..30]
  
  unless workbook.sheet_by_name(sheet_name).nil?
    repeated_sheet_name_number = 2
    until workbook.sheet_by_name(sheet_name).nil? do
      # Add (#) to the end of the sheet name. Name must be 31 or less chars, so
      # the end of some function names may need to get replaced by the suffix.
      end_of_shortened_name = 30 - (repeated_sheet_name_number.to_s.length+2)
      if end_of_shortened_name < 2
        puts "Warning: Failed to write #{starting_title} sheet to Excel file.".red
        return
      end
      
      sheet_name = sheet_name[0..end_of_shortened_name] + "(#{repeated_sheet_name_number})"
      repeated_sheet_name_number += 1
    end
    puts "Worksheet name for #{starting_title} is: #{sheet_name}"
  end
  
  sheet_name
end