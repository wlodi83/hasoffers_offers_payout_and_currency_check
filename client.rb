require File.join(File.dirname(__FILE__), 'requester')
require 'spreadsheet'
require 'yaml'
require 'yajl'

#Load HasOffers config details
@config = YAML.load_file("config.yaml")
@hasoffers_api_url = @config["config"]["api_url"]
@network_id = @config["config"]["network_id"]
@network_token = @config["config"]["network_token"]

#Create xls result file
@result = Spreadsheet::Workbook.new
@sheet1 = @result.create_worksheet
@sheet1.name = 'Payout Comparison between HasOffers and AMS'
@sheet1.row(0).concat %w{LandingPageID HasOffersID ENABLED AMS_VALUE AMS_CURRENCY HO_PAYOUT HO_CURRENCY}

row_counter = 1

#Read xls file with information from SP database about landing pages
#Read Excel File
Spreadsheet.open('landing_pages.xls') do |book|
  book.worksheet('Sheet1').each do |row|
    break if row[0].nil?
      next if row[1] == "AFFILIATE_OFFER_ID"
        response = Requester.make_request(
        @hasoffers_api_url,
        {
          "Format" => "json",
          "Service" => "HasOffers",
          "Version" => "2",
          "NetworkId" => "#{@network_id}",
          "NetworkToken" => "#{@network_token}",
          "Target" => "Offer",
          "Method" => "findById",
          "id" => "#{row[1]}"
        },
        :get
      )

      json = StringIO.new("#{response}")
      parser = Yajl::Parser.new
      hash = parser.parse(json)
      if hash["response"]["data"].nil?
        excel_row = @sheet1.row(row_counter)
        excel_row[0] = row[0]
        excel_row[1] = row[1]
        excel_row[2] = row[2]
        excel_row[3] = row[3]
        excel_row[4] = row[4]
        excel_row[5] = "No data"
        excle_row[6] = "No data"
      else

        payout = hash["response"]["data"]["Offer"]["default_payout"]
        currency = hash["response"]["data"]["Offer"]["currency"]

        if payout.to_f != row[3].to_f && currency != row[4]
          excel_row = @sheet1.row(row_counter)
          excel_row[0] = row[0]
          excel_row[1] = row[1]
          excel_row[2] = row[2]
          excel_row[3] = row[3]
          excel_row[4] = row[4]
          excel_row[5] = hash["response"]["data"]["Offer"]["default_payout"]
          excel_row[6] = hash["response"]["data"]["Offer"]["currency"]
        end

      end
      row_counter += 1
  end
end

@result.write 'result.xls'
