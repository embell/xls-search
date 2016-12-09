xls-search Documentation
========================

Description
-----------
This utility uses Ruby libraries to parse xls spreadsheets for a list of keywords and then generate an Excel report from the results. 

This was made and tested on Windows. It was also made for personal convenience; not much consideration for widespread usefulness was involved, so your mileage may vary.

Setup
-----
To use, you need Ruby installed. 

Once you have Ruby, follow these steps for setup:
1. Clone the project folder using git clone.
2. Run _bundle install_ inside the folder if using Bundler. If not, run _gem install [name]_ for each gem in the Gemfile.

How to use
----------
1. Copy the xls-search folder to your preferred folder.
2. Verify that the *settings.txt* file points to the folder containing spreadsheets to search.
3. Open the *search_terms.txt* file and enter the keywords to search for, one per line.
4. Run *main.rb*. Once execution is complete, a report should open up in Excel. The report is saved to the xls-search/temp folder.
