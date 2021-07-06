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
our_messages = []

for message in messages_retrieved_from_junkfolder:
    if(message.subject == "Interested Lead"): #and message.created.day == todays_day):
        message_body = message.get_body_text()
        print(message_body)
        for line in message_body:
            with open("Todays Leads.txt", "w+") as leads_file:
                leads_file.writeline("Leads")
                leads_file.write(message_body)
        print("Message Date: ", message.created.day)
        #print(message_body)
            



"""
for i in our_messages:
   print(i.body)
"""




""" Sending Leads """
try:
    my_message = mailbox.new_message()
    my_message.to.add(["my_email"])
    my_message.subject = "FWD: Lead"
    my_message.body = "Lead\n"
    my_message.attachments.add("Todays Leads.txt")
    my_message.send()
    print("Message Sent Successfully")
except Exception:
    with open("log.txt", 'w+') as logfile:
        logfile.write(traceback.print_exc())
    print("Log File Saved Successfully")

