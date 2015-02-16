require 'mail'

class SmtpMail 
	def initialize
		@mail = Mail.new do
		  from    'lwang74@live.cn'
		  # from    'wang.lin@inventec.com'
		  # to      'wang.lin@inventec.com'
		  # subject 'Mail from Mail'
		  # body    'There is a body.'
		end
	end

	def deliver to, subject, html_body, attach_file=nil
		@mail.to = to
		@mail.subject = subject
		# @mail.body = html_body

		text_html = Mail::Part.new do
  			body "<h1>ruby mail text/html</h1>"
		end

		@mail.html_part = text_html

		# file_data = File.read(attach_file)
		# @mail.attachments['abc.htm'] = {:content => file_data} if attach_file

# :mime_type => 'application/x-pdf', 
		@mail.delivery_method :smtp, { address: 'smtp.live.com', #address:   'tao-cs1.iec.inventec'}
		                              port:      587, authentication: :tls}
		                           #    domain:    'itc.inventec',
		                           #    user_name: 'itc940167',
		                           #    password:  'naomi_94a167',
		                           #    authentication: :login,
		                       		  # enable_starttls_auto: true}
		                              # password:  ($stderr.print 'password> '; gets.chomp) }
		@mail.deliver!
	end
end

def send_mail to, subject, html_body, attach_file=nil
	my_mail = SmtpMail.new
	my_mail.deliver to, subject, html_body, attach_file
end

if __FILE__==$0
	send_mail 'wang.lin@inventec.com', 'Subject ssdfd', "HTML Context <font style='color:red'>ABCD</font>", 'F:\82_lwang\Salary\李玉香_2013年08月.htm'
end