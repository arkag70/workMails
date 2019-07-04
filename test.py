import win32com.client as win32
import pandas as pd


def getDataFrame():

	data = pd.read_excel("team.xlsx", sheet_name="Sheet1")
	name = ""
	with open("file1.txt","w") as f:
		for i in range(len(data)):
			name += data.iloc[i]["Name"]+", "
		f.write(name)

def mail(to,sub,body):
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	mail.To = to # separated by semi colon
	mail.Subject = sub
	mail.Body = body
	mail.HTMLBody = f'<h2>{mail.body}</h2>' #this field is optional

	# To attach a file to the email (optional):

	mail.Send()

if __name__ == "__main__":
	getDataFrame()