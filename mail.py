import os.path
import openpyxl
from operator import is_not
from functools import partial
from string import Template
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from email.mime.base import MIMEBase




def get_contacts(filename):
       wb = openpyxl.load_workbook(os.path.join('#add path to file', filename))
       sheet = wb.get_sheet_by_name('#name of the sheet')
       names = []
       emails = []
       for row in range(2, sheet.max_row + 1):
              emails.append(sheet['B' + str(row)].value)
              names.append(sheet['A' + str(row)].value)

       names = filter(partial(is_not, None), names)
       emails = filter(partial(is_not, None), emails)
       names = [str(item) for item in names]

       return names , emails

def read_template(filename):
    filepath = os.path.join('#add path to file', filename)

    with open(filepath, 'r') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)



email_username = 'example@gmail.com'   #add you email username
email_password = 'example'  # add your email password



def main():
    names, emails = get_contacts('example.xlsx') # read contacts
    message_template = read_template('invite.txt')

    # set up the SMTP server
    s = smtplib.SMTP(host='imap.gmail.com', port=587)  #if sending mail through gmail
    s.starttls()
    s.login(email_username, email_password)

    # For each contact, send the email:
    count = 0
    while True:
        for name, email in zip(names, emails):
               msg = MIMEMultipart()  # create a message

               # add in the actual person name to the message template
               message = message_template.substitute(PERSON_NAME=name.title())

               # Prints out the message body for our sake
               print(message)

               # setup the parameters of the message
               msg['From'] = email_username
               msg['To'] = email
               msg['Subject'] = ""   #Subject of Email

               msg.attach(MIMEText(message, 'plain'))

               f = '#add attachment'
               part = MIMEBase('application', "octet-stream")
               part.set_payload(open(f, "rb").read())
               encoders.encode_base64(part)
               part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
               msg.attach(part)
              
               # send the message via the server set up earlier.
               s.send_message(msg)

               del msg

        # Terminate the SMTP session and close the connection

        s.quit()


if __name__ == '__main__':
    main()
