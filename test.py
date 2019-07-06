import win32com.client as win32
import pandas as pd
from datetime import datetime


def getDataFrame():

	to = ""
	cc = ""
	data = pd.read_excel("team.xlsx", sheet_name="Sheet1")
	todaysdate = str(datetime.now()).split(' ')[0]
	time = str(datetime.now()).split(' ')[1]
	thismonth = int(todaysdate.split('-')[1])
	thisday = int(todaysdate.split('-')[2])
	#print(f"thisday:{thisday} and thismonth:{thismonth}")
	
	for entry in range(len(data)):
		birthdate = str(data.iloc[entry]["DOB"]).split(" ")[0]
		birthmonth = int(birthdate.split("-")[1])
		birthday = int(birthdate.split("-")[2])

		if(thismonth == birthmonth and thisday == birthday):
			to += data.iloc[entry]["Name"]+"; "
		else:
			cc += data.iloc[entry]["Mail_ID"]+"; "
	names = " ".join([x for x in to.split("; ")])
	sub = f"Happy Birthday to : {names}"
	print(f"To : {to}")
	print(f"CC : {cc}")
	print(f"Sub: {sub}")
	 
	
	

def mail(to,cc,sub,body):
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	mail.To = to # separated by semi colon
	mail.CC = cc
	mail.Subject = sub
	mail.Body = body
	mail.HTMLBody = f'<h2>{mail.body}</h2>' #this field is optional

	# To attach a file to the email (optional):

	mail.Send()

if __name__ == "__main__":
	getDataFrame()