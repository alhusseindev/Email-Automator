from O365 import Account, FileSystemTokenBackend, message, connection, MSGraphProtocol
import datetime
import traceback

#Authentication

credentials = ("Client ID", "Client Secret")
tenant_id = "Tenant ID"

account = Account(credentials, auth_flow_type='credentials', tenant_id= tenant_id)

if(account.authenticate()):
    print('Authenticated!')
else:
    print("UnAuthenticated / Authentication Error!")



todays_date = datetime.datetime.now()
todays_day = todays_date.day
todays_month = todays_date.month
todays_year = todays_date.year

#Accessing mailbox

mailbox = account.mailbox("Email")

inbox = mailbox.inbox_folder()
sent_folder = mailbox.sent_folder()
junk_folder = mailbox.junk_folder()
messages_retrieved_from_inbox = inbox.get_messages()
messages_retrieved_from_sentfolder = sent_folder.get_messages()
messages_retrieved_from_junkfolder = junk_folder.get_messages(limit= 5, download_attachments= True)




#Taking care of messages
blacklisted_countries = ["abc", "def", "ghi"] #Countries With High Number of Spammers!
blacklisted_keywords = ["abcSpam", "defSpam", "ghiSpam"] #keywords that spammers use
skip_message = False

with open("Leads.txt", "a+", encoding="UTF-8") as leads_file:
    for message in messages_retrieved_from_junkfolder:
        if(message.subject == "Some Subject" and message.created.day == todays_day):
            print("Found Email That I am Interested In!")
            message_body = message.get_body_text()
	    message_body.lower()
            for country in blacklisted_countries:
                if(country in message_body):
                    print("BlackListed Country! Skip this inquiry. ", "Country: " + country)
                    skip_message = True
                    #break out of the country loop
                    break
               
            for keyword in blacklisted_keywords:
                if(keyword in message_body):
                    print("Blacklisted Keyword! Skip this inquiry.", "Keyword: " + keyword)
                    skip_message = True
                    #break out of the keyword loop
                    break

            if(not skip_message):
                leads_file.write("--------------------------------------------------\nLead:\n--------------------------------------------------\n")
                leads_file.write(message_body)

            

        else:
            print("Not Interested in Email!")
            
                


#checking the file
with open("Leads.txt", 'r') as file_being_read:
    for line in file_being_read:
        if(len(line) == 0):
            print("Leads File Is Empty, Will Not Send Email!")
    
    """ Sending Leads """
    try:
        my_message = mailbox.new_message()
        my_message.to.add(["email@email.com"])
        my_message.subject = "FWD: Lead(s)"
        my_message.body = "Lead(s)\n"
        my_message.attachments.add("Leads.txt")
        my_message.send()
        print("Message Sent Successfully")
    except Exception:
        logging.basicConfig(filename= "email_error.log", filemode= "w", level= logging.DEBUG) 
        print("Log File Saved Successfully")


            
#Deleting the file after it's being sent
os.remove("Leads.txt")
print("Leads File Deleted Successfully")



