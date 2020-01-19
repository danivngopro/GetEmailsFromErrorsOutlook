import email
import os

path = 'C:\\Users\\Daniel\\Desktop\\files\\'
listing = os.listdir(path)
f = open("BadEmails.txt", "a")
for fle in listing:
    msg = email.message_from_file(open(path+fle))
    msg2 = str(msg).split("Final-recipient:")[1]
    msg2 = msg2.split("Action:")[0]
    msg2 = msg2.split("; ")[1]
    print(msg2)
    f.write(msg2)
f.close()