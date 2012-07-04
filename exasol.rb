require 'rubygems'
require 'dbi'

class Exasol

  def initialize(login, password)
    @login = login
    @password = password
  end

  def connect
    begin
      #connect to the Exasol Server
      @dbh = DBI.connect('dbi:ODBC:EXA', @login, @password)
    rescue DBI::DatabaseError => e
      puts "An error occured"
      puts "Error code: #{e.err}"
      puts "Error message: #{e.errstr}"
      err_message = "#{e.errstr}"
    end
  end

  def disconnect
    #disconnect from server
    @dbh.disconnect if @dbh
  end

  def do_query(query)
    q = query
    begin
      @sth = @dbh.prepare(q)
      @sth.execute
    rescue DBI::DatabaseError => e
      puts "An error occured"
      puts "Error code: #{e.err}"
      puts "Error message: #{e.errstr}"
      err_message = "#{e.errstr}"
    end
  end

  def print_result_array
    @result_array = Array.new
    while row=@sth.fetch_array do
      @result_array.push(row)
    end
    @sth.finish
    return @result_array
  end

end
