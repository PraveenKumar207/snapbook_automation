#!/usr/bin/env ruby
require 'net/http'
require 'uri'
require 'base64'
require 'open-uri'
require 'json'
require 'byebug'
require 'rubyXL'
require 'rubyXL/convenience_methods'


# Static things that needs to be changed are cohort_id and img_url
def img_url
    @img_url.gsub(' ', '')
end

def content
    Base64.encode64(open(img_url).to_a.join)
end

def cohort_id
    @cohort_id.to_i
end

def uri
    URI::HTTPS.build(
        host: 'slapi-staging.tllms.com', 
        path: '/goggles-qs/gogglesFind',
        query: URI.encode_www_form({cohort_id: cohort_id})
    )
end

def request
    req = Net::HTTP::Post.new(uri).tap do |req|
        req["X-TNL-USER-ID"] = 2700246
        req["X-TNL-TOKEN"] = 'TGnH4ZoDCJv7BAnqFTHjYZyF'
        req["X-TNL-APPID"] = 1
        req["Content-Type"] = 'application/json'
      end
    req.body = {content: content }.to_json
    req
end

def get_response
    response = Net::HTTP.start(uri.host, uri.port, use_ssl: uri.scheme == "https",
                                                                open_timeout: 5, read_timeout: 20) do |http|
            http.request(request)
        end

    res = JSON.parse(response.body)
end

file = ARGV[0]
workbook = RubyXL::Parser.parse(file)
worksheet = workbook[0]

@headers = worksheet[0].cells.map{|c| c.nil? ? nil : c.value}

key_col = @headers.find_index('key')
img_col = @headers.find_index('preview_image_url')
cohort_col = @headers.find_index('cohort_id')
result_page_key_col = cohort_col + 2
results_found_col = cohort_col + 3
high_accurate_result_col = cohort_col + 4
high_accurate_page_col = cohort_col + 5

worksheet.add_cell(0, result_page_key_col, 'result_page_key')
worksheet.add_cell(0, results_found_col, 'results_found?')
worksheet.add_cell(0, high_accurate_result_col, 'high_accurate_result?')
worksheet.add_cell(0, high_accurate_page_col, 'high_accurate_page')


@last_index = 0
@match_count = 0
@total_count = 0
worksheet.each_with_index do |row, row_index|
    next if !row || row_index == 0 || worksheet[row_index][img_col]&.value.nil? || worksheet[row_index][cohort_col]&.value.nil?
    row.cells.each_with_index do |cell, col_index|
        @img_url = cell.value if col_index == img_col
        @cohort_id = cell.value if col_index == cohort_col
    end
    begin 
        response = get_response
        puts "\nResponse for row #{row_index} => \n#{response}"

        #result_page_key
        result_page_key = response['metadata']['pid']
        worksheet.add_cell(row_index, result_page_key_col, result_page_key)
        if result_page_key != worksheet[row_index][key_col].value

            worksheet.sheet_data[row_index][result_page_key_col].change_fill('ff337e')
        else
            @match_count = @match_count + 1
        end
        @total_count = @total_count + 1

        #results_found
        results = response['results']
        worksheet.add_cell(row_index, results_found_col, (!results.empty?).to_s)

        #high_accurate_result
        high_accurate_result_found = !response['metadata']['accurate']['pid'].to_s.empty? rescue 'false'
        worksheet.add_cell(row_index, high_accurate_result_col, high_accurate_result_found.to_s)

        #high_accurate_page
        high_accurate_page = response['metadata']['accurate']['pid'] || 'Not Found'
        worksheet.add_cell(row_index, high_accurate_page_col, high_accurate_page)
    rescue => e
        worksheet.add_cell(row_index, result_page_key_col, "An Exception occured")
    end
    
    @last_index = row_index
end

accuracy_percent = @match_count.to_f*100/@total_count.to_f

result_percent_row = @last_index + 2
result_percent_col = result_page_key_col
worksheet.add_cell(result_percent_row, result_percent_col, 'accuracy %')
worksheet.add_cell(result_percent_row, result_percent_col + 1, accuracy_percent)


workbook.write(file)

puts "\n -- Completed --" 
