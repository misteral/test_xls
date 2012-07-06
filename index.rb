# encoding: utf-8

require 'spreadsheet'

file_r = ENV['HOME']+"/projects/NOT_NOW/new_xls/3.xls"
file_s = ENV['HOME']+"/projects/NOT_NOW/new_xls/4.xls"



book = Spreadsheet.open file_r
 sheet = book.worksheet 0
 sheet.each do |row|
   row[0] = "Тестируем"
 end
 sheet[5,5]="Произвольное втыкание"

 book.write file_s


