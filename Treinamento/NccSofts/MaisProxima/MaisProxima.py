import smtplib

def email(email_subject, email_body, email_to_email):
    email_from_email = 'workflow@maisproxima.com.br'
    email_from_user = 'workflow@maisproxima.com.br'
    email_from_password = 'MP#qwert1234' 

    smtp = smtplib.SMTP(host = 'smtp.office365.com', port = 587)
    smtp.ehlo()
    smtp.starttls()
    smtp.ehlo

    smtp.login(email_from_user, email_from_password)
    email_header = \
        'To:' + ', '.join(email_to_email) + '\n' + \
        'From: ' + email_from_email + '\n' + \
        'Subject: ' +  email_subject + '\n\n'
    smtp.sendmail(email_from_email, email_to_email, email_header + email_body)
    smtp.close()
    
    return
