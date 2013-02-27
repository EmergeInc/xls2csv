# coding:utf-8
require 'spreadsheet'
require 'csv'

module Xls2Csv
  class Core
    def initialize(xls, dir, logger = nil)
      @xls = xls
      @dir = dir
      @logger = logger || Logger.new(STDOUT)
      Spreadsheet.client_encoding = 'UTF-8'
    end

    def start
      #Xls2Csv::Options
      convert
    end

    def convert
      read_xls.each {|sheet_name, value| write_csv(sheet_name, value)}
    end

    def read_xls(xls = @xls)
      if xls.respond_to?(:read) # File-like object
        raise Ole::Storage::FormatError unless File::extname(xls.original_filename).downcase == '.xls'
      else
        raise Ole::Storage::FormatError unless File::extname(xls).downcase == '.xls'
      end

      csvs = Hash.new{|hash, key| hash[key] = []}
      Spreadsheet.open(xls).worksheets.each do |sheet|
        sheet.each {|row| csvs[sheet.name] << row_to_s(row)}
      end
      csvs
    rescue Errno::ENOENT, Errno::EACCES
      @logger.info 'Reading ERROR!!!'
      @logger.info $!.message
    rescue Ole::Storage::FormatError
      @logger.info 'Reading ERROR!!!'
      @logger.info "#{xls} is not xls-file."
    end

    def write_csv(filename, value)
      File.open("#{@dir}/#{filename}.csv", 'w') do |f|
        value.each {|row| f.puts "\"#{row.join('","')}\""}
      end
    rescue Errno::ENOENT, Errno::EACCES
      @logger.info 'Writing ERRER!!!'
      @logger.info $!.message
    end

    private
    def row_to_s(row)
      #binding.pry
      row
      #row.map do |cell|
        #if cell.class == Spreadsheet::Formula
          #cell.value
        ##elsif cell.class == Fixnum
          ##row.date(0).strftime('%m/%d/%Y')
        #else
          #cell.to_s
        #end
      #end
    end
  end
end
