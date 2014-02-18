require './c_excelcls.rb'

module Trim
	def trim str
		if str
			str.gsub(/[ 　]/, '')
			# str.gsub("\241\241", "")
		else
			str
		end
	end
end

class CMailExcel
	attr :all
	attr :info
	include Excel
	include Trim

	def my_range rg
		result_hash ={}
		is_title = true
		rg.value2.each{|row|
			break if !row[0] or row[0].to_s.strip == ''
			#~ if !is_title and row[2].to_s.strip != ''
			if !is_title
				#~ result_hash[trim(row[0])] = [trim(row[1]), row[2].to_s.strip!='']
				yield result_hash, row
			else
				is_title = false
			end
		}
		result_hash
	end

	def initialize book
		open(book){|wb|
			@all = my_range(wb.worksheets('MailAddress').usedrange){|h, row|
				name_alias = nil
				if row[4]
					row[4].strip!
					name_alias = row[4].split(',')
				end

				h[trim(row[0])] = [trim(row[1]), row[2].to_s.strip!='', row[3], name_alias]
			}
			@info = my_range(wb.worksheets('info').usedrange){|h, row|
				h[trim(row[0])] = row[1]
			}
		}
		# p @all
	end
  
	def get_mail_add_fname person_name
		per = @all[person_name]
		if per[1]
			if per[2] && per[2].strip != ''
				[per[0], "<span style='color:blue'>#{per[2]}</span>"] #mail address, full name.
			else
				[per[0], "<span style='color:blue'>#{person_name}</span> 老师"] #mail address, full name.
			end
		else
			nil
		end
	end
  
	def get_right_name name_or_alias
		right_name = name_or_alias
		@all.each{|name, val|
			if name== name_or_alias
				right_name = name
				break
			else
				if val[3]
					val[3].each{|name_alias|
						if name_alias== name_or_alias
							#~ puts "#{name_alias} => #{name}"
							right_name = name
							break
						end
					}
				end
			end
		}
		right_name
	end
	
	def chk_name_exists person_name
		#~ @all.include?(person_name)
		ret = false
		@all.each{|name, val|
			if name== person_name
				ret = true
				break
			else
				if val[3]
					val[3].each{|name_alias|
						if name_alias== person_name
							#~ puts name_alias
							ret = true
							break
						end
					}
				end
			end
		}
		ret
	end
end

if __FILE__==$0
	total_dir = "#{Dir.getwd.gsub("/", "\\")}\\total"
	$mail_addr = CMailExcel.new("#{total_dir}\\教师邮箱.xls")
	p $mail_addr.all
	p $mail_addr.info
end
