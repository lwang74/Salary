# coding: utf-8
require 'FileUtils'
require './c_excelcls'
require './mailexcel'
require './send_mail'

class CSalaryExcel
	attr :all
	attr :month
	include Excel
	include Trim
  
	def initialize total_dir, info
		book = "#{total_dir}\\#{info['input']}"
		@month = info['month']
		@desc = info['desc']
		@all={}
		@sheetname = []
    
		#~ Here
		@sheet_info = []
		@person_info = {}
		@fail_person = {}
		
		open(book){|wb|
			sheet_index = 0
			wb.worksheets.each{|x|
			    #~ p x.name
			    @all[x.name] = {}
			    @sheet_info[sheet_index] = {:name=>x.name}
			    sheet_proc(wb.WorkSheets(x.name).usedrange, @all[x.name], x.name, sheet_index)
			    #~ @sheetname << x.name
			    sheet_index += 1
			}
		}
	end
  
	def check log_file
		File.open(log_file, "w"){|out|
			@fail_person.each{|name, dumy|
				out.puts name
			}
		}
	end
  
	def out_people_htm file_out_path=nil
		file_out_path = @month if !file_out_path
		begin
			FileUtils.mkdir file_out_path
		rescue
		end
		@person_info.each{|k, v|
			File.open("#{file_out_path}\\#{k}_#{@month}.htm", "w"){|out|
				out.puts out_person_htm{|context|
					context = out_person_htm_table(k, v)
				}
			}
		}
	end
  
  	def get_real_body name, fname
  		@desc.gsub(/\[name\]/, name).gsub(/\[fname\]/, fname).gsub(/\[month\]/, @month)
  	end

	def send_each_mail realy_send, check_htm
		chk = out_person_htm{|chk_out|
			chk_all_table = ""
			cnt = 0
			first = true
			@person_info.each{|k, v|
				#~ mail_addr = $mail_addr.get_mail_add(k)
				mail_addr, fname = $mail_addr.get_mail_add_fname(k)

				if mail_addr
					# mail_content =  out_person_htm{|out|
					# 	out = out_person_htm_table(k, v)
					# }
					chk_all_table += out_person_htm_bady(k, fname) if first
					first=false
					chk_all_table += out_person_htm_table(k, v)
					if realy_send
						atta_file = "#{FileUtils.pwd.gsub("/", "\\")}\\#{@month}\\#{k}_#{@month}.htm"
						subject = "*** #{k} *** 天津八十二中学#{@month}工资明细表"
						#~ send_mail mail_addr, "*** #{k} *** 天津八十二中学#{@month}份工资明细表", mail_content, atta_file
						send_mail mail_addr, subject, get_real_body(k, fname), atta_file
					end
				end
				cnt += 1
			}
			chk_out = chk_all_table
		}

#~ p check_htm
		File.open(check_htm, "w"){|out|
			out.puts chk
		}
	end
	
	protected
	def sheet_proc rg, hash, sheet_name, sheet_index
		#~ rows = []
		#~ is_title = true
		hash['person'] ||= {}
		hash['title'] = []
		name = nil
		is_header = true
		header = []
		header_add = {}
		name_cnt = 0
		name_add = {}
		rg.rows.each{|row|
			if is_header
				header_row = []
				row.columns.each{|cell|
					if cell.MergeCells
						if !header_add[cell.MergeArea.Address]
							header_row << merge_cell(cell)
							header_add[cell.MergeArea.Address] = true
						else
							header_row << nil
						end
					else
						header_row << cell.Text
					end

					#~ puts cell.text+cell.Address
					if '姓名'==trim(cell.text)
						if cell.MergeCells
							name_add[cell.MergeArea.Address] = true
							name_cnt += 1
							#~ puts "name_cnt=#{name_cnt}; Add: #{cell.MergeArea.Address}"
						else
							is_header = false
						end
					else
						# puts cell.text
						if cell.MergeCells
							if name_add[cell.MergeArea.Address]
								name_cnt += 1
								#~ puts "name_cnt=#{name_cnt}; Add: #{cell.MergeArea.Address}"
								is_header = false if name_cnt==cell.MergeArea.FormulaLocal.size
							end
						end
					end
				} 
				header<<header_row
			else
				row_arr = []
				name = ''
				row.columns.each{|cell|
					name = $mail_addr.get_right_name(trim(cell.text)) if 2==cell.column
					row_arr << cell_text(cell)
				}
				if ''!=name
					@fail_person[name] = '' if !$mail_addr.chk_name_exists(name)
					@person_info[name] ||= []
					@person_info[name][sheet_index] = row_arr
				end
			end
		}
		@sheet_info[sheet_index][:header] = header
		if is_header
			@fail_person["发现一个Sheet格式有错误，请检查。Sheet名：'#{sheet_name}'"]=''
		end
	end
	
	def merge_cell cell
		[cell.Text, cell.MergeArea.FormulaLocal.size, cell.MergeArea.FormulaLocal[0].size]
	end
	
	def cell_text cell
		if cell.Text=~/#+/
			cell.Value2
		else
			cell.Text
		end
	end
	
	def sheet_html_out header, sheet_name, context, index
		#~ p header
		out = ""
		out += "<table border='0' cellpadding='0' cellspacing='0' class='outline'>\n"
		out += "<tr><td class='bgc#{index%5}' style='color:white;width:150px;text-align:center'>#{sheet_name}</td><td style='width:90%'>&nbsp;</td></tr>\n"
		out += "<tr><td colspan='2'><table border='0' cellpadding='0' cellspacing='0' class='inner#{index%5}'>\n"
		header.each_with_index{|row, row_index|
			row_str = []
			out += '<tr>'
			row.each_with_index{|col, col_index|
				left = (0==col_index)? '' : 'left'
				if col.class == Array
					#~ row_str << "#{col[0]}->#{col[1]}->#{col[2]}"
					row_str << "<th class='#{left} bottom' rowspan='#{col[1]}' colspan='#{col[2]}'>#{col[0]}</th>"
				else
					row_str << "<th class='#{left} bottom'>#{(''!=col)? col : '&nbsp;'}</th>" if col
				end
			}
			out += row_str.join("")
			out += "</tr>\n"
		}
		out += "<tr><td>#{context.map{|x| if !x or x=='' then '&nbsp;' else x end}.join("</td><td class='left'>")}</td></tr>\n"
		out += "</table></td></tr>\n"
		out += "</table>\n"
	end

	def out_person_htm_bady name, fname
		out = "<table style='border:solid 1px gray;' cellspacing='0' cellpadding='0'>\n"
		out += "<tr><td>#{get_real_body(name, fname)}</td></tr>\n"
		out += "</table>"
	end

	def out_person_htm_table name, content
		out = "<table style='border:solid 1px gray;' cellspacing='0' cellpadding='0'>\n"
		out += "<tr class='tr_name'><td colspan='2' class='td_name'>天津八十二中学#{@month}工资明细表</td></tr>\n"
		out += "<tr class='tr_name'><td colspan='2' class='td_name td_border'><span style='font-size:26px'>#{name}</span></td></tr>\n"
		cnt = 0
		@person_info[name].each{|sht|
			if sht
				out += "<tr><td colspan='2'>\n"
				out += sheet_html_out(@sheet_info[cnt][:header], @sheet_info[cnt][:name], sht, cnt)+"\n"
				out += "</td></tr>"
			end
			cnt +=1
		}
		out += "</table>"
	end
  
	def out_person_htm
		colors = ['rgb(192,80,77)', 'rgb(75,172,198)', 'rgb(247,150,70)', 'rgb(155,187,89)', 'rgb(128,100,162)']
		out = <<-END_OF_STRING
    <html>
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" /> 
    <title>工资条</title>
    <style>
	.tr_name {color:white;background-color:rgb(31,73,125);}
	.td_name {padding-left:50px; color:white}
	.td_border {border-left:solid 1px gray; border-top:solid 1px gray;}
	
	.bgc0 {color:white; background-color:#{colors[0]};}
	.bgc1 {color:white; background-color:#{colors[1]};}
	.bgc2 {color:white; background-color:#{colors[2]};}
	.bgc3 {color:white; background-color:#{colors[3]};}
	.bgc4 {color:white; background-color:#{colors[4]};}
	
	.bd0 {solid 1px #{colors[0]}}
	.bd1 {solid 1px #{colors[1]}}
	.bd2 {solid 1px #{colors[2]}}
	.bd3 {solid 1px #{colors[3]}}
	.bd4 {solid 1px #{colors[4]}}
	
	table {border: solid 1px gray}
	table.outline {margin:20px; border: solid 0px gray}
	
	table.inner0 {border: solid 2px #{colors[0]};}
	table.inner1 {border: solid 2px #{colors[1]};}
	table.inner2 {border: solid 2px #{colors[2]};}
	table.inner3 {border: solid 2px #{colors[3]};}
	table.inner4 {border: solid 2px #{colors[4]};}
 	
	th.left {border-left: solid 1px gray;}
	th.bottom {border-bottom: solid 1px gray;}
	td {color:blue;}
	td.left {border-left: solid 1px gray;}
	td.bottom {border-bottom: solid 1px gray;}
    </style>
    </head>
    <body>
END_OF_STRING
		out1 = nil
		out += yield(out1)
		out += "</body></html>"
	end
end

def main
	total_dir = "#{Dir.getwd.gsub("/", "\\")}\\total"
	$mail_addr = CMailExcel.new("#{total_dir}\\教师邮箱.xls")

	salary = CSalaryExcel.new(total_dir, $mail_addr.info)
	log_file = "error.txt"
	salary.check log_file
	if !File::zero?(log_file)
		system("notepad #{log_file}")
	else
		salary.out_people_htm

		chk_htm = "c:\\check.htm"
		salary.send_each_mail false, chk_htm
		ie = WIN32OLE.new('InternetExplorer.Application')
		ie.visible = true
		ie.navigate "file:///#{chk_htm}"
		print "Generate mail?(y/n)"
		STDOUT.flush
		ret = gets
		if ret=~/y/i
			salary.send_each_mail true, chk_htm
		end
	end
end

main

# For test below
# 	total_dir = "#{Dir.getwd.gsub("/", "\\")}\\total"
# 	$mail_addr = CMailExcel.new("#{total_dir}\\教师邮箱.xls")
# $mail_addr.all.each{|k, v|
# 	p v
# }
