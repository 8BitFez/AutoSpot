from copy import error
import imap_tools
import csv
import configparser

# -------------------------------------------------
#
# Utility to read email from Gmail Using Python
#
# ------------------------------------------------
try:
    config = configparser.ConfigParser()
    config.read('config.ini')


    SCAN_EMAIL = config['GMAIL']['EMAIL']
    FROM_PWD = config['GMAIL']['PASSWORD']
    SMTP_SERVER = config['SMTP']["SMTP_SERVER"]
    SMTP_PORT = config['SMTP']['SMTP_PORT']
    FROM_EMAIL = config['GMAIL']['FROM_EMAIL']
except error:
    print("config error\n")
    print(error)
    input("\n press enter to continue")

DATE_STR = "%Y%M%D"
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
            key = (date_str + msg.subject)
            # create the values and add it to the database
            for keyword in FIND_WORDS:
                out = find_value(msg.text,keyword)
                arr = out.split("* ")
                value_out[arr[0]] = arr[1]

            value_out[UUID] = key 
            DATABASE.append(value_out)
    
    print(str(DATABASE))
    
    with open(CSV_FILE, mode='w') as csv_file:
        fieldnames = FIND_WORDS
        fieldnames.append(UUID)
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

        writer.writeheader()
        for email in DATABASE:
            writer.writerow(email)

if __name__ == "__main__":
    try:
        
        main()
    except error:
        print(error)
        input("\n Press enter to quit")
    else:
        print("Clean exit i think?")
        input("\n Press enter to quit")