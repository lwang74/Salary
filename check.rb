File.open("names.txt", "w"){|out|
  File.new("uri.txt").each{|x|
    if x=~/\*\*\*\s+(.+)\s+\*\*\*/
      out.puts $1
    end
  }
}
