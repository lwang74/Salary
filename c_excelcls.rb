#~ Last modified: 2008-5-15

require 'win32ole'
# p WIN32OLE.codepage
WIN32OLE.codepage = WIN32OLE::CP_UTF8

module Excel
	def open book, save_it=false
		excel = WIN32OLE.new('Excel.Application')
		#~ excel.visible = TRUE
		workbook = excel.Workbooks.open(book)
		yield workbook
    workbook.save if save_it
		workbook.close true
		#~ excel.quit
	end

	def rows start
		tm = start
		tm_val = tm.text.to_s.strip
		ret = []
		while tm_val!=''
			ret.push tm_val
			tm = tm.offset(0, 1)
			tm_val = tm.text.to_s.strip
		end
		ret
	end

	def trim_nil arr
		tr =true
		arr_new = arr.delete_if{|x| #delet before nil
			tr=false if x
			tr
		}
		tr =true
		arr_new_new = arr_new.reverse.delete_if{|x|#delete after nil
			tr=false if x
			tr
		}
		arr_new_new.reverse
	end
end

module Arr_strip
  def strip!
    self.each{|x|
      x =x.strip if x and x.class==String
      self[self.index(x)] = x.round.to_s if x.class==Float
    }
  end
end

#~ Building Spec
class CSpec
	attr :all
	include Excel
	
  def range rg
		ret={}
		sec = nil
		rows = []
		rg.value2.each{|row|
      row.extend Arr_strip
      row.strip!
			if row[0]
				break if row[0].to_s.upcase == 'END'
				sec = row[0]
				rows = ret[sec.downcase] = []
			elsif row[1] and sec
				rows << trim_nil(row)
			else
				sec = nil
			end
		}
		ret
	end
  
	def initialize book, prj_id=nil
		@all={}
		open(book){|wb|
      if prj_id #for individule project
        @all = range(wb.WorkSheets(prj_id).usedrange)
        range(wb.WorkSheets('common').usedrange).each{|k,v|
          @all[k] = v if !@all.has_key?(k) 
        }
      else
        wb.worksheets.each{|x|
          @all[x.name] = range(x.usedrange)
        }
      end
		}
	end
end

#~ Drivers Excel file 'driver list' sheet.
class CDrvlist
	attr :all
	include Excel
	def initialize book
		open(book){|wb|
			drvsht = wb.worksheets('driver list')
			@all=range(drvsht.usedrange)
		}
		#~ p @all
	end

protected
	def range rg
		ret={}
		rg.value2.each{|row|
			if row[0]
        row.extend Arr_strip
        row.strip!
				break if row[0].to_s.upcase == 'END'
				row_new =  trim_nil(row)
#~ p row_new[0].class==Float
				if Float==row_new[0].class
					row_new[0] = row_new[0].to_i.to_s
					#~ p row_new[0]
				end
				ret[row_new.shift] = row_new
			end
		}
		ret
	end
end

#~ Drivers Excel file 'ID' sheet.
class CIDs
	attr :drv_type
	attr :os_list
	attr :ven_model
	include Excel
  def initialize book
    @drv_type = []
    @os_list = {}
    @ven_model = []
    
		open(book){|wb|
			drvsht = wb.worksheets('ID')
			range(drvsht.usedrange)
		}
  end
  
protected
  def range rg
		ret={}
    cnt = 0
		rg.value2.each{|row|
      break if !row[0] and !row[1] and !row[2] and !row[3]
			if cnt>0
        @drv_type << row[0] if row[0]
        @os_list[row[1]]=row[2] if row[1]
        @ven_model << row[3] if row[3]
      end
      cnt += 1
		}
	end
end

if __FILE__==$0
  t1 = Time.new
  
class CExcel
	attr :all
	attr :info
	include Excel
	def initialize book
		open(book){|wb|
			wb.worksheets.each{|sht|
				p sht.name
				p sht.name.encoding
			}
		}
	end
end

excel = CExcel.new("F:\\82_lwang\\Salary\\total\\教师邮箱.xls")

  puts "time #{Time.new - t1} is spent."
end


