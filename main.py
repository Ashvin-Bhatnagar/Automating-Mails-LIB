import gspread
import smtplib
import ssl
from email.message import EmailMessage

# initializing all the prerequisites to open the sheet and send mails
SHEET_ID = '1NpdDp18cE7ywBf9ZBd4D5VeXMDbVdDXNolCUvLcIqr0'
EMAIL_PASSWORD = "vlwamwvjgdcnqial"
EMAIL_SENDER = "ashvin.bhatnagar@oberoi-is.net"
EMAIL_SENDER_ADDRESS = "<ashvin.bhatnagar@oberoi-is.net>"

# subject and body template remain the same for each mail
subject = "Let it Beat Order Receipt"
body = """Hello, from everyone here at Let it Beat, thanks so much for ordering! 
Here is your receipt, and please note that the last date of payment is the 20th of January. 
"""

# connecting to Google's port and setting up the email, logging in to the sender
email = EmailMessage()
context = ssl.create_default_context()
smtp = smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context)
smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)

# opening the spreadsheet
gc = gspread.service_account("automatinglibmails-7433b83275db.json")
spreadsheet = gc.open_by_key(SHEET_ID)

# setting various fields of the email
email["From"] = EMAIL_SENDER + " " + EMAIL_SENDER_ADDRESS
email["Subject"] = subject
email.set_content(body)

# automating emails
# iterating through each grade (each sheet)
for i in range(1, 7):
    # opening each sheet and getting list of mails
    worksheet = spreadsheet.get_worksheet(i)
    mails = worksheet.col_values(2)

    # trimming list of mails to exlude headers, if there are any mails
    if len(mails) != 0:
         for x in range(3):
             mails.pop(x)

    # sending a mail to everyone on the list
    for mail in mails:

        # formatting receiver as per Google RFC-5321 protocols
        to = "<" + mail + ">"
        email["To"] = to

        # sending mails to people who have not paid and have not received a mail yet (new responses)
        if worksheet.cell((mails.index(mail)+1), 36).value != "DONE":
            if worksheet.cell((mails.index(mail)+1), 8).value != "TRUE":

                # ensuring that any incorrect emails are not included (avoiding errors)
                if to != "<>" and ("@" in mail):
                    smtp.sendmail(EMAIL_SENDER_ADDRESS, to, email.as_string())

                    # marking every already sent-to mail address as done
                    worksheet.update_cell(mails.index(mail)+1, 36, "DONE")

        # resetting the "to" field so there is only one recipient in each mail
        del email["To"]
