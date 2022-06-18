from copy import error
from logging import exception
import imap_tools
import csv
import configparser

# -------------------------------------------------
#
# Utility to parse SpotHero Email to database
#
# ------------------------------------------------
error_code = 0
config_err = False
config = configparser.ConfigParser()
config.read('config.ini')

if(config.has_section('GMAIL')):
    SCAN_EMAIL = config['GMAIL']['EMAIL']
    FROM_PWD = config['GMAIL']['PASSWORD']
    FROM_EMAIL = config['GMAIL']['FROM_EMAIL']
else:
    config_err = True
if(config.has_section('SMTP')):
    SMTP_SERVER = config['SMTP']["SMTP_SERVER"]
    SMTP_PORT = config['SMTP']['SMTP_PORT']
else:
    config_err = True
if(config_err):
    error_code = 1
    print("\nConfig File error please check your config file\n")


DATE_STR = "%Y/%m/%d"
CSV_FILE = "Database.csv"
# KeyWords to fine in email 
FIND_WORDS = ["Rental ID#:","Reservation Start:","Reservation End:","License Plate:","Phone Number:"]
UUID = "name"
#Keywords with The value
DATABASE = []

#Find and snip out the import stuff
def find_value(body,find_str):
    start = body.find(find_str)
    body_lang = len(body)
    for x in range(start,body_lang):
        scn = body[x]
        end = x
        if scn == "\n":
            sub = body[start:end]
            return sub
            break
def main():
    with imap_tools.MailBox(SMTP_SERVER).login(SCAN_EMAIL, FROM_PWD) as mailbox:
        for msg in mailbox.fetch(imap_tools.AND(from_= FROM_EMAIL)):
            value_out = {}
            #Create a unique key for every reservation 
            date_str =msg.date.strftime(DATE_STR)
            key = (date_str +"_"+ msg.subject)
            # create the values and add it to the database
            for keyword in FIND_WORDS:
                out = find_value(msg.text,keyword)
                arr = out.split("* ")
                value_out[arr[0]] = arr[1]

            value_out[UUID] = key 
            DATABASE.append(value_out)
    for email in DATABASE:
        print('added '+str(email[UUID]))
        print("added "+str(len(DATABASE))+" reservation to database")
    
    with open(CSV_FILE, mode='w') as csv_file:
        fieldnames = FIND_WORDS
        fieldnames.append(UUID)
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

        writer.writeheader()
        for email in DATABASE:
            writer.writerow(email)

if __name__ == "__main__":
    try:
        if(error_code > 0):
            raise error
        main()
    except Exception as e:
        print("error_code:" + str(e))
        input("Press enter to quit")
    else:
        print("Clean exit error code:"+ str(error_code))
        input("\nPress enter to quit")