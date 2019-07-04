import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'Arkaprabha.ghosh@in.bosch.com;Menon.Shekhar@in.bosch.com'
mail.Subject = 'Hello'
mail.Body = 'Hello testing from python'
mail.HTMLBody = f'<h2>{mail.body}</h2>' #this field is optional

# To attach a file to the email (optional):


mail.Send()