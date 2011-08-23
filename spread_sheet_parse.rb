#!/usr/bin/env ruby

require'rubygems'
require'roo'
require 'spreadsheet'


class Parsesheet
	STARTCOLUMN=7
	STARTROW=5
	def initialize in_file,out_file
		@week_count=0
		@output=[]
		begin
			@outputfile=out_file
			@sheet=Openoffice.new(in_file)
			@sheet.default_sheet=@sheet.sheets.first
		rescue
			puts "Cannot open the spreadsheet file #{filename}"
		end
	end
	
	def check_no_weeks
		x=STARTCOLUMN+1
		while @sheet.cell(1,x)==nil
			@week_count+=1
			x+=1
		end
		@week_count
	end
	
	def find_column_indexes
		@colindexes=[]
		x=STARTCOLUMN
		@colindexes<<x
		while x < @sheet.last_column
			x+=@week_count+1
			@colindexes.push(x)
		end
		@colindexes.pop
		return @colindexes
	end
	
	def find_row_indexes
		@rowindexes=[]
		c=0
		x=STARTROW
		while c!=2
			unless pro=@sheet.cell(x,1)
				c+=1
				x+=1
				next
			end
			if !((pro.downcase == "totals" )|| (pro.downcase.include? "internal"))
				@rowindexes.push(x)
			end
			x+=1
		end
		@rowindexes
	end
	
	def calculate_gross_work_time
		record=[]
		hash={}
		@output=[]
		sum=0
		@rowindexes.each do |row|
			record.push(@sheet.cell(row,1))     # pushing the project name
			@colindexes.each do |col|
				(1..@week_count).each { |wk| sum+=@sheet.cell(row,col+wk-1).to_f }
				if sum >0
					hash[@sheet.cell(1,col)]=sum
				end
				sum=0
			end
			record.push(hash)                   #pushing the code worktime details of each person
			hash={}
			@output.push(record)
			record=[]
		end
		@output
	end
	
	def write_to_sheet 
		Spreadsheet.client_encoding='UTF-8'
		book=Spreadsheet::Workbook.new
		sheet1=book.create_worksheet
		sheet1[0,0]="Project Name"
		sheet1[0,1]="Resource Name"
		sheet1[0,2]="Number of hours"
		row=1
        format = Spreadsheet::Format.new :color => :red
                                 
		@output.each do |project|	
		 sheet1[row,0]=project[0].encode("UTF-8")
		 sheet1.row(row).default_format=format
			row+=1
       		 project[1].each do |programmer| 
				sheet1[row,1]=programmer[0].encode("UTF-8")
				sheet1[row,2]=programmer[1]
				row+=1
			end		
		end
		book.write(@outputfile)
	end

end

if ARGV.length!=2 
	puts "Usage: spread_sheet_parse <inputfile> <outputfile> \n eg ./spread_sheet_parse a.ods b.xls"
	exit
end

sh=Parsesheet.new(ARGV[0],ARGV[1])  # put your file name here
sh.check_no_weeks
sh.find_column_indexes
sh.find_row_indexes
sh.calculate_gross_work_time
sh.write_to_sheet 
