from O365 import Account, FileSystemTokenBackend, message, connection, MSGraphProtocol
import datetime
import traceback
import logging
import os




#Authentication

credentials = ("3c7ed49e-9b84-4f53-b9c7-754c7efca86d", "7Z~KWqDcP3P9~B670YlWmwlEf.g-SQ1_eh")
tenant_id = "46c18b75-a1e3-49af-a41f-2c9725c590f5"

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

mailbox = account.mailbox("info@wellaliments.com")

inbox = mailbox.inbox_folder()
sent_folder = mailbox.sent_folder()
junk_folder = mailbox.junk_folder()
messages_retrieved_from_inbox = inbox.get_messages()
messages_retrieved_from_sentfolder = sent_folder.get_messages()
messages_retrieved_from_junkfolder = junk_folder.get_messages(limit= 30, download_attachments= True)



#Taking care of messages
blacklisted_countries = ["benin", "Benin","Togo", "togo", "South Africa", "south africa", "Ukraine", "taiwan", "Taiwan", "Afghanistan", "afghanistan", "ukraine", "Uruguay", "uruguay","Nigeria", "Netherlands", "netherlands","Fiji", "fiji", "Poland", "poland", "Ghana", "ghana", "switzerland", "Switzerland", "India", "China", "china", "Pakistan", "pakistan", "Kenya", "kenya", "Bangladesh", "bangladesh"]
blacklisted_keywords = ["Packaging", "packaging", "database", "Database", "Email List", "email list",  "interia.pl", "soap", "Soap", "cash", "Cash", "online", "Online", "shop", "Shop", "list", "List","seo", "SEO", "PPC", "ppc", "mail.ru"]
skip_message = False

with open("Leads.txt", "a+", encoding="UTF-8") as leads_file:
    for message in messages_retrieved_from_junkfolder:
        if(message.subject == "Wellaliments - Contact us" and message.created.day == todays_day):
            print("Found Well Aliments Email!")
            message_body = message.get_body_text()
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
            print("Not Well Aliments Email!")
            
                


#checking the file
with open("Leads.txt", 'r') as file_being_read:
    for line in file_being_read:
        if(len(line) == 0):
            print("Leads File Is Empty, Will Not Send Email!")
        else:
            """ Sending Leads """

            try:
                my_message = mailbox.new_message()
                my_message.to.add(["aeweis08@gmail.com"])
                my_message.subject = "FWD: Lead(s)"
                my_message.body = "Lead(s)\n"
                my_message.attachments.add("Leads.txt")
                my_message.send()
                print("Message Sent Successfully")
            except Exception:
                logging.basicConfig(filename= "email_error.log", filemode= "w", level= logging.DEBUG) 
                print("Log File Saved Successfully")



os.remove("Leads.txt")
print("Leads File Deleted Successfully")


