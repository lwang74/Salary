require 'mail'

mail = Mail.new do
  from    'wang.lin@inventec.com'
  to      'wang.lin@inventec.com'
  subject 'Mail from Mail'
  body    'There is a body.'
end

mail.delivery_method :smtp, { address:   'tao-cs1.iec.inventec'}
                           #    port:      25,
                           #    domain:    'itc.inventec',
                           #    user_name: 'itc940167',
                           #    password:  'naomi_94a167',
                           #    authentication: :login,
                       		  # enable_starttls_auto: true}
                              # password:  ($stderr.print 'password> '; gets.chomp) }
mail.deliver!
