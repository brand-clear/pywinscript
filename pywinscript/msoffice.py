

import win32com.client as win32
from pywintypes import com_error


def send_email(to, cc, subject, body, open_email=False):
	"""Send email via MS Outlook.

	Parameters
	----------
	to : list[str]
	cc : list[str]
	subject : str
	body : str
	open_email : {True, False}, optional

	Raises
	------
	com_error
        Generally raised for network-related issues.

	Returns
	-------
	True
        Email was sent.
		
	Notes
	-----
	An open instance of MS Outlook is assumed.

	"""
	try:
		outlook = win32.Dispatch("Outlook.Application")
		email = outlook.CreateItem(0)
		email.To = "; ".join(to)
		email.Cc = "; ".join(cc)
		email.Subject = subject
		email.GetInspector
		index = email.HTMLbody.find(">", email.HTMLbody.find("<body")) 
		email.HTMLbody = email.HTMLbody[:index+1]+body+email.HTMLbody[index+1:] 
		if open_email:
			email.Display(open_email)
		else:
			email.Send()
	except com_error:
		raise
	else:
		return True



# TEST
# send_email(
# 	['brandon.mccleary@sulzer.com'],
# 	['brandon.mccleary@sulzer.com'], 
# 	'Title', 
# 	'Body'
# )
