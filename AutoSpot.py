
import imap_tools
import csv
import configparser
import datetime
import numpy as np
# -------------------------------------------------
#
# Utility to parse SpotHero Email to database
#
# ------------------------------------------------
error_code = 0
config_err = False
CSV_FILE = "Database.csv"

config = configparser.ConfigParser()
config.read('config.ini')
TODAY = datetime.date.today()
## checking if the config has these section
##
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
FIND_WORDS = ["Rental ID#:","Reservation Start:","Reservation End:","License Plate:","Phone Number:"]
UUID = 'Uid'
DATE = 'Date'
FIELD_NAMES = np.concatenate((FIND_WORDS,[UUID,DATE]))


print("Today is "+(TODAY.strftime(DATE_STR)))
# KeyWords to fine in email
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
            sub = body[start:(end)]
            sub_str =  sub.strip()
            return sub_str
            break    
def get_from_sender(Database,From_Email,imported=False,Sent_Date=None):
    with imap_tools.MailBox(SMTP_SERVER).login(SCAN_EMAIL, FROM_PWD) as mailbox:
        for msg in mailbox.fetch(imap_tools.AND(from_= From_Email,sent_date=Sent_Date)):
            #Create a unique key for every reservation 
            value_out = {UUID:msg.uid,DATE:msg.date}
            dup = False
            # create the values and add it to the database
            for keyword in FIND_WORDS:
                out = find_value(msg.text,keyword)
                arr = out.split("* ")
                value_out[arr[0]] = arr[1]
#         dup check 
            if imported:    
                for data in Database:
                    if value_out[UUID] == data[UUID]:
                        print ("found duplicate entry skipping UID: " + data[UUID])
                        dup = True
            if not dup:
                Database.append(value_out)
                print("added email with UID: "+value_out[UUID])  
def write_to_csv(database):
    with open(CSV_FILE, mode='w') as csv_file:
        fieldnames = FIND_WORDS
        fieldnames.append(UUID,)
        fieldnames.append(DATE)
        
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

        writer.writeheader()
        for email in database:
            writer.writerow(email)
def read_csv(database):
    success = True
    try:
        with open(CSV_FILE) as csv_file:
            csv_reader = csv.DictReader(csv_file, delimiter=',')
            for row in csv_reader:
                database.append(row)
    except KeyError as e:
        print("Import error could not find key " + str(e))
        print("Entry could have deleted/changed")
        success = False
    except FileNotFoundError as e:
        print("File not found or Corrupted can't import ")
        success = False
    finally:
        return success
def main():
    import_suc = read_csv(DATABASE)
    get_from_sender(DATABASE,FROM_EMAIL,imported=import_suc)
    print(str(DATABASE))
    write_to_csv(DATABASE)

if __name__ == "__main__":
    main()