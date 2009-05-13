from imaplib import *
from poplib import *
import string, StringIO, rfc822

#####################################################################################
#
# email_transfer.py, by Carl Mercier (carl@carlmercier.com)
#
# This script will transfer emails from one server (POP or IMAP) to another (IMAP)
# with the option of deleting the messages on the source server.
#
# For IMAP source servers, only the INBOX will be transfered.
#
# Version 1.0 - January 31, 2007
#
#####################################################################################


# -----------------------------------------------------------------------------------
# Leave the rest as is!
# -----------------------------------------------------------------------------------

# A few helper methods...
def login_pop3(host,port,username,password):
  server = POP3(host,port)
  server.getwelcome()
  server.user(username)
  server.pass_(password)
  return server

def login_imap4(host,port,username,password):
  server = IMAP4(host,port)
  server.login(username, password)
  return server

def get_message_list(server_type, server):
  if server_type == "IMAP":
    resp, items = server.search(None, "ALL")
    items = string.split(items[0])
    return items
    
  elif server_type == "POP":
    items = server.list()[1]
    i = 0
    for item in items:
      items[i] = item.split(' ')[0]
      i +=1
    return items
    
  else:
    raise "Invalid SOURCE_SERVER_TYPE."

def retrieve_message(server_type, server, msg_id):
  if server_type == "IMAP":
    text = server.fetch(msg_id, "(RFC822)")[1][0][1]
  elif server_type == "POP":
    text = "\n".join(server.retr(msg_id)[1])
  else:
    raise "Invalid SOURCE_SERVER_TYPE."
  file = StringIO.StringIO(text)
  message = rfc822.Message(file)
  return message, text
  
def get_date_from_message(message):
  return rfc822.parsedate(message['date'])

def delete_message(server_type, server, msg_id):
  if server_type == "IMAP":
    server.store(msg_id, '+FLAGS', '\\Deleted')
    server.expunge()
  elif server_type == "POP":
    server.dele(msg_id)
  else:
    raise "Invalid SOURCE_SERVER_TYPE."

  
# Do it.
for mailbox in MAILBOXES:
  mailbox = mailbox.replace(",",":")
  tmp = mailbox.split(":")

  print "Processing user '" + tmp[0] + SOURCE_SERVER_USERNAME_SUFFIX + "'..."
  
  source_server_user = tmp[0] + SOURCE_SERVER_USERNAME_SUFFIX
  source_server_password = tmp[1]
  
  if len(tmp) > 2:
    target_server_user = tmp[2] + TARGET_SERVER_USERNAME_SUFFIX
    target_server_password = tmp[3]
  else:
    target_server_user = tmp[0] + TARGET_SERVER_USERNAME_SUFFIX
    target_server_password = tmp[1]
   
  print "Connecting to source server: " + SOURCE_SERVER_HOST + "..."
  if SOURCE_SERVER_TYPE == "IMAP":
    source_server = login_imap4(SOURCE_SERVER_HOST, SOURCE_SERVER_PORT, source_server_user, source_server_password)
  elif SOURCE_SERVER_TYPE == "POP":
    source_server = login_pop3(SOURCE_SERVER_HOST, SOURCE_SERVER_PORT, source_server_user, source_server_password)
  else: 
    raise "Invalid SOURCE_SERVER_TYPE."
      
  print "Connected."
  
  print "Connecting to target server: " + TARGET_SERVER_HOST + "..."
  target_server = login_imap4(TARGET_SERVER_HOST, TARGET_SERVER_PORT, target_server_user, target_server_password)
  print "Connected."

  # go to INBOX
  if SOURCE_SERVER_TYPE == "IMAP":
    source_server.select()
  
  target_server.select()
  
  print "Listing items on source server..."
  items = get_message_list(SOURCE_SERVER_TYPE, source_server)
  
  counter = 0
  
  for msg_id in items:
    counter += 1
    print ""
    print "-------------------------------------------------------------------------------"
    print "Processing message " + str(counter) + "/" + str(len(items))
    message, text = retrieve_message(SOURCE_SERVER_TYPE, source_server, msg_id)
    msg_date = get_date_from_message(message)
    try:
      sender = message['from']
    except:
      sender = ""
    
    try:
      subject = message['subject']
    except:
      subject = ""
    
    print "FROM: " + sender + ", SUBJECT: " + subject
    print "Uploading to target server..."
    target_server.append("INBOX", None, msg_date, text)
    
    if DELETE_MESSAGES_FROM_SOURCE_SERVER:
      print "Deleting from source server..."
      delete_message(SOURCE_SERVER_TYPE, source_server, msg_id)
      
    print "-------------------------------------------------------------------------------"
    print ""
  
  print "Disconnecting..."
  if SOURCE_SERVER_TYPE == "IMAP":
    source_server.logout()
  elif SOURCE_SERVER_TYPE == "POP":
    source_server.quit()
    
  target_server.logout()

print "All done!"
