require 'win32ole'
require 'singleton'

class Outlook 
  include Singleton
  def initialize 
    @ol = WIN32OLE::connect("Outlook.Application")
    WIN32OLE.const_load(@ol,self.class)
  end

  def new_mail
    mail =  @ol.CreateItem(OlMailItem)
    #~ mail.BodyFormat = olFormatHTML
    return mail
  end
end

def send_mail to, subject, html_body, attach_file=nil
  outlook = Outlook.instance
  mail = outlook.new_mail
  #~ mail.To = to
  to.split(",").each{|one_to|
    mail.Recipients.Add one_to
  }
  mail.Subject = subject
  mail.HTMLBody = html_body
  mail.Attachments.Add(attach_file) if attach_file
  #~ mail.GetInspector.Activate
  mail.Send
end

if __FILE__==$0
  send_mail 'rubima@cuzic.com,test@abc.com', 'WIN32OLE test mail', "WIN32OLE <font style='color:red'>ABCD</font>"
end