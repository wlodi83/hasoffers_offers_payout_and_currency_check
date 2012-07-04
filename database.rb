require File.join(File.dirname(__FILE__), 'exasol')
require 'spreadsheet'
require 'yaml'

config = YAML.load_file("config.yaml")
@login = config["config"]["login"]
@password = config["config"]["password"]

#Create result file
result_excel = Spreadsheet::Workbook.new
sheet1 = result_excel.create_worksheet
sheet1.name = 'Result'
sheet1.row(0).concat %w{ProgramName	CountryCode	LandingPageID	HasOffersID	ACTIVE_IN_ALL_LEVELS IGNORE_COMMUNICATED_PAYOUT_ENABLED?	AMS_VALUE	AMS_CURRENCY HO_PAYOUT	HO_CURRENCY	Payout Mismatch}

row_counter = 1

@connection = Exasol.new(@login, @password)
@connection.connect

Spreadsheet.open('ignore.xls') do |book|
  book.worksheet('Result').each do |row|
    next if row[2] == "LandingPageID"
    lp = row[2]

    puts lp

    query_1 = "select lp.id, prov.ignore_communicated_provision \
               from cms.advertisers as adv \
               join cms.programs as prog on adv.id = prog.advertiser_id \
               join cms.program_regions as pr on prog.id = pr.program_id \
               join cms.countries as co on pr.country_id = co.id \
               join cms.landing_pages as lp on lp.program_region_id = pr.id \
               join cms.provisions as prov on prov.landing_page_id = lp.id \
               where lp.\"enabled\" = 1 and adv.\"enabled\" = 1 and prog.\"enabled\" = 1 and pr.\"enabled\" = 1 and lp.id = '#{lp}'"

    @connection.do_query(query_1)
    result_1 = @connection.print_result_array
    puts result_1
        if !result_1.empty?
          excel_row = sheet1.row(row_counter)
          excel_row[0] = row[0]
          excel_row[1] = row[1]
          excel_row[2] = row[2]
          excel_row[3] = row[3]
          excel_row[4] = "Yes"
          excel_row[5] = result_1[0][1]
          excel_row[6] = row[5]
          excel_row[7] = row[6]
          excel_row[8] = row[7]
          excel_row[9] = row[8]
          excel_row[10] = row[9]
          excel_row[11] = row[10]
        end

      row_counter += 1

  end

end

@connection.disconnect
result_excel.write 'final_result.xls'
