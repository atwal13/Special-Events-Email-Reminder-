import smtplib
import openpyxl

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


my_address = "" #my address here 
pw = "" #fill in the PW here 



def main():
    wb = openpyxl.load_workbook(filename = "Special Events.xlsx")
    sheet1 = wb["Sheet1"]
    message = "%-50s %-30s %-50s" % ("Name of Event", "Date of Event", "Location of Event")
    message += "\n"

  


    for row in sheet1.iter_rows():
        if row[0].value == "Name":
            continue

        message += "%-50s %-30s %-50s" % (row[0].value, row[1].value, row[2].value)
        message += "\n\n"
   
    print(message)
    s = smtplib.SMTP(host = "smtp-mail.outlook.com", port = 587)
    s.starttls()
    s.login(my_address, pw)
    print("Logged in!")

    msg = MIMEMultipart()


    msg["From"] = my_address
    msg["To"] = "" #Sending it to whom?
    msg["Subject"] = "Special Events Reminder "

    message += "The list above shows the major special events from multiple sources online."
    message += "\nThanks for reading\n\n Best Regards, \n\n"
    
    msg.attach(MIMEText(message, "plain"))

    s.send_message(msg)
    print("Sent!")

    del msg

    s.quit()

if __name__ == "__main__":
    main()
    print("Done")
