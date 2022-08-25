# snapbook_automation

## Setup:
1. Install ruby by `brew install ruby`
2. Install rvm and install ruby 2.6.6. Refer [this](https://nrogap.medium.com/install-rvm-in-macos-step-by-step-d3b3c236953b) write up.
3. gem install 'net/http'
4. gem install 'uri'
5. gem install 'base64'
6. gem install 'open-uri'
7. gem install 'json'
8. gem install 'byebug'
9. gem install 'rubyXL'

## Steps to use the script:
1. Copy the code in automation_script.rb file and save it in a file (lets say automation.rb)
2. Download the excel sheet (lets say sample_sheet.xlsx)
3. Command to execute: `ruby  automation.rb  sample_sheet.xlsx` (ruby [space] path/to/script.rb [space] path/to/excel_file.xlsx)

Sample sheet https://docs.google.com/spreadsheets/d/1Qf69GDqofBND3t3w-7VUnZC3080t353-OJtHMiA45tQ/edit#gid=648119026
