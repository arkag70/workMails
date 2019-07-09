import win32com.client as win32
import pandas as pd
from datetime import datetime
import os
import random

def html_src(names):
	
	files = []
	files = os.listdir("\\\\bosch.com\\dfsrb\\DfsIN\\LOC\\Kor\\NE2\\ESV_Info\\36_ESX\\Common\\General_Information\\DASy\\09_TeamDetails\\ESX3BirthdayList\\Cards\\")
	path = "\\\\bosch.com\\dfsrb\\DfsIN\\LOC\\Kor\\NE2\\ESV_Info\\36_ESX\\Common\\General_Information\\DASy\\09_TeamDetails\\ESX3BirthdayList\\Cards\\"+str(random.choice(files))

	html_content = f'''

	<!DOCTYPE html>
	<html>
		<head> 
		</head>
		<body>
			<h2 style ="font-family:'Comic Sans MS', cursive, sans-serif; color:#005cde; font-size:48px">Many many happy returns of the Day {names}</h2>
				<img src = {path} ></img>

		</body>
	</html>'''
	return html_content

def getDataFrame():

	to = ""
	cc = ""
	names = ""
	data = pd.read_excel("\\\\bosch.com\\dfsrb\\DfsIN\\LOC\\Kor\\NE2\\ESV_Info\\36_ESX\\Common\\General_Information\\DASy\\09_TeamDetails\\ESX3BirthdayList\\team.xlsx", sheet_name="Sheet1")
	todaysdate = str(datetime.now()).split(' ')[0]
	time = str(datetime.now()).split(' ')[1]
	thismonth = int(todaysdate.split('-')[1])
	thisday = int(todaysdate.split('-')[2])
	#print(f"thisday:{thisday} and thismonth:{thismonth}")
	
	for entry in range(len(data)):
		birthdate = str(data.iloc[entry]["DOB"]).split(" ")[0]
		birthmonth = int(birthdate.split("-")[1])
		birthday = int(birthdate.split("-")[2])
		# name = str(data.iloc[entry]["Name"])
		# print(f"{name}  date : {birthday} month : {birthmonth}")

		if(thismonth == birthmonth and thisday == birthday):
			to += str(data.iloc[entry]["Mail_ID"])+"; "
			names += str(data.iloc[entry]["Name"])+"; "
		else:
			if "bosch.com" in str(data.iloc[entry]["Mail_ID"]):
				cc += str(data.iloc[entry]["Mail_ID"])+"; "

	names = " and ".join([x for x in names.split("; ")]) 
	names = names[:-4]
	sub = f"Happy Birthday {names}"
	body = html_src(names)
	mail_func(to,"",sub,body)
	
	 
	
	

def mail_func(to,cc,sub,body):
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	mail.To = to # separated by semi colon
	mail.CC = ""
	mail.Subject = sub
	mail.Body = body
	mail.HTMLBody = body #this field is optional

	# To attach a file to the email (optional):

	mail.Send()

if __name__ == "__main__":
	getDataFrame()
	#Arkaprabha.ghosh@in.bosch.com
