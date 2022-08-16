#!/usr/bin/env ruby
require 'net/http'
require 'uri'
require 'base64'
require 'open-uri'
require 'json'
require 'byebug'
require 'rubyXL'

# Static things that needs to be changed are cohort_id and img_url
def img_url
    @img_url
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

img_col = @headers.find_index('preview_image_url')
cohort_col = @headers.find_index('cohort_id')
results_found_col = cohort_col + 2
high_accurate_result_col = cohort_col + 3
high_accurate_page_col = cohort_col + 4

worksheet.add_cell(0, results_found_col, 'results_found?')
worksheet.add_cell(0, high_accurate_result_col, 'high_accurate_result?')
worksheet.add_cell(0, high_accurate_page_col, 'high_accurate_page')

worksheet.each_with_index do |row, row_index|
    next if row_index == 0 || worksheet[row_index][img_col].nil? || worksheet[row_index][cohort_col].nil?
    row.cells.each_with_index do |cell, col_index|
        @img_url = cell.value if col_index == img_col
        @cohort_id = cell.value if col_index == cohort_col
    end
    response = get_response
    puts "\nResponse for row #{row_index} => \n#{response}"
    results = response['results']
    worksheet.add_cell(row_index, results_found_col, (!results.empty?).to_s)

    #high_accurate_result
    high_accurate_result_found = !response['metadata']['accurate']['pid'].to_s.empty? rescue 'false'
    worksheet.add_cell(row_index, high_accurate_result_col, high_accurate_result_found.to_s)

    #high_accurate_page
    high_accurate_page = response['metadata']['accurate']['pid'] rescue 'Not Found'
    worksheet.add_cell(row_index, high_accurate_page_col, high_accurate_page)
    
end

workbook.write(file)

puts "\n -- Completed --" 



