## Venezuela Automation ##

Application to automate Venezuela Production.

- Using Google API (gmail), to extract Stock Exchange email information from gmail account:  automation.marketdatalatam@gmail.com
- Using Google API (spreadsheet), as Front-End to manage Holiday, Recipients and see log of sent emails.
	https://docs.google.com/spreadsheets/d/1yGgknknCMacg7ifpFYejQzFjXAcYpdKuUzwiulkqaNM/edit?usp=sharing

Dependences: 
	- credentials.json: Yuo must registry your project on 'https://console.developers.google.com/' to be able
			    create Client IDs OAuth 2.0, necessary to access gmail email.

	- service_account: Yuo must registry your project on 'https://console.developers.google.com/' to be able
			   create Service Account, necessary to access google sheets.

	- token.pickle:  If there isn't, it make automatically after you give permission via browser.
Run:


	> pip install -r requirements.txt
	> python start.py