"""
 *****************************************************************************
   FILE :           contact.py

   AUTHORS :        Jiin Jeong and Heather Wing

   DATE :           June 6 - June 8, 2018 (Completed)
                    July 24, 2018 (Improved & Cleaned)

   DESCRIPTION :    Sends email to a list of addresses in an excel file
                    that contains weather information.
                    (Sensitive info removed.)
   IMPROVEMENTS:
   (1) Style
   (2) Email class & main function
   (3) Able to send e-mails to multiple ppl.

 *****************************************************************************
"""

import smtplib  # Library for email transmission
import xlrd  # Reads Exel
from weather import Weather, Unit  # Yahoo Weather info
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


class Email():
    """ Configures SMTP e-mail settings. """

    def __init__(self, account, domain, pw, to, msg):

        self._account = account
        self._server = smtplib.SMTP_SSL(domain, 465)  # Standard port
        self._server.ehlo()  # Optional, called by login()
        self._pw = pw
        self._to = to
        self._msg = msg

    def sendEmail(self):
        """ Sends e-mail or prints error message. """

        try:
            # Logs in and sends message.
            self._server.login(self._account, self._pw)
            self._server.sendmail(self._account, self._to, self._msg)
            self._server.close()
            print('successfully sent the mail')

        except:
            print("Something went wrong...")


class ExcelData():
    """ Reads in Excel data. """

    def __init__(self, sheet):

        self._sheet = sheet

    def readSheet(self, column, firstRow, lastRow):
        """ Reads Excel and returns the email address in the row. """
        
        emailList = []

        # Adjusts row start/end values to match the excel sheet's labels.
        start = firstRow -1
        end = lastRow

        for rowID in range(start, end):
            row = self._sheet.row(rowID)
            emailList.append(row[column-1].value)
        
        print("Distribution list: ", emailList)
        return emailList


def main():
    """ Loads data and sends e-mail. """

    # Reads in contact list sheet.
    workbook = xlrd.open_workbook('contact_list.xlsx')
    sheet = ExcelData(workbook.sheet_by_index(0))
    to = sheet.readSheet(2, 2, 3)  # Col, startRow, endRow
    name = sheet.readSheet(1, 2, 3)
    # print(to)  # Debugging purposes.
    # print(name)

    # Set up sender email. Change.
    account = 'senderemail'
    domain = "smtp.gmail.com"
    pw = 'senderpw'

    # Loads weather data in Fahrenheit.
    weather = Weather(unit=Unit.FAHRENHEIT)

    # Looks up weather for the location (Claremont) using WOEID.
    loc = weather.lookup(2380633)
    cur_temp = loc.condition.temp

    # Set up message.
    msg = MIMEMultipart()
    msg['From'] = account
    msg['To'] = "UTC Distribution List"
    msg['Subject'] = "Remember to Water Your Tree!"
    msg.preamble = 'Sustainable Claremont'

    # Attach test image.
    fp = open('contact_img.png', 'rb')
    img = MIMEImage(fp.read())
    fp.close()
    msg.attach(img)

    # Sends e-mail to every recipient on the list.
    for i in range(len(to)):
        body = ("Hello %s! \n\n" % (name[i]) + 
                "The weather today is %s degrees. " % (cur_temp) +
                "This can put your tree at risk for deydration! " +
                "Go out and water yours today.\n\n" +
                "Thank you from your friends at Sustainable Claremont\n")
        msg.attach(MIMEText(body, 'plain'))

        email = Email(account, domain, pw, to, msg.as_string())
        email.sendEmail()

if __name__ == '__main__':
  main()
