import datetime
import json
import smtplib
import traceback

import parametros

# GMail : Less secure app access
def email(email_subject, email_body, email_to_email):
    email_from_email = parametros.global_email_from_email
    email_from_user = parametros.global_email_from_user
    email_from_password = parametros.global_email_from_password
    email_smtp_host = parametros.global_email_smtp_host
    email_smtp_port = parametros.global_email_smtp_port

    smtp = smtplib.SMTP(host=email_smtp_host, port=email_smtp_port)
    smtp.ehlo()
    smtp.starttls()
    smtp.ehlo

    smtp.login(email_from_user, email_from_password)
    email_header = \
        "To:" + ", ".join(email_to_email) + "\n" + \
        "From: " + email_from_email + "\n" + \
        "Subject: " +  email_subject + "\n\n"
    smtp.sendmail(email_from_email, email_to_email, email_header + email_body)
    smtp.close()
    
    return

def ok(header, is_email=True):
    if is_email:
        email_header = ("{:%Y-%m-%d %H:%M:%S} " + header + " OK") \
            .format(datetime.datetime.now())
        email_body = ""
        email_to_email = parametros.global_email_to_email
        email(email_header, email_body, email_to_email)

    return

def excecao(exception, header, is_email=True):
    print("{:%Y-%m-%d %H:%M:%S} {:}\n{:}"
          .format(datetime.datetime.now(), str(exception), traceback.format_exc()))
          #.format(datetime.datetime.now(), sys.exc_info()[0].__doc__))
    
    if is_email:
        email_header = ("{:%Y-%m-%d %H:%M:%S} " + header) \
            .format(datetime.datetime.now())
        email_body = "{:}\n{:}" \
            .format(str(exception), traceback.format_exc())
            #.format(sys.exc_info()[0].__doc__, traceback.format_exc())
        email_to_email = parametros.global_email_to_email
        email(email_header, email_body, email_to_email)

    return

def print_class(x):
    y = json.dumps(x.__dict__, indent=4)
    #y = json.dumps(x.__dict__, indent=4, sort_keys=True)
    print(y)

def print_json(x):
    y = json.load(x)
    z = json.dumps(x, indent=4)
    #z = json.dumps(y, indent=4, sort_keys=True)
    print(z)

def print_object(x):
    y = json.dumps(x, indent=4)
    #y = json.dumps(x, indent=4, sort_keys=True)
    print(y)
