import email
import os
import win32com.client
import extract_msg
from independentsoft.msg import Message
import codecs

f = codecs.open("BadEmailsDoctors", "w", "utf-8")
path = 'C:\\Users\\itmanager.BMR\\Desktop\\emailstocheck\\'
listing = os.listdir(path)
for fle in listing:
	appointment = Message(path + fle)
	x = str(appointment.display_to)
	if(".com" in x):
		x = x.split(".com")[0]
		x = x + ".com"
	if(".il" in x):
		x = x.split(".il")[0]
		x = x + ".il"
	if(".net" in x):
		x = x.split(".net")[0]
		x = x + ".net"
	if(".ru" in x):
		x = x.split(".ru")[0]
		x = x + ".ru"
	if(".uk" in x):
		x = x.split(".uk")[0]
		x = x + ".uk"
	if(".eu" in x):
		x = x.split(".eu")[0]
		x = x + ".eu"
	if(".fr" in x):
		x = x.split(".fr")[0]
		x = x + ".fr"
	if("@gmail" in x):
		x = x.split("@gmail")[0]
		x = x + "@gmail.com"
	f.write(x)
	f.write("\n")
f.close()