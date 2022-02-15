# Outlook template creator
A program for searching for a lease channel using Excel tables and preparing a template for sending a letter.

Important:
the program works with an Excel file containing information about the channels. The search for channels is carried out by columns with the name of the equipment at the connection points. The received information contains email addresses for communication. An Outlook template is then created for the received email address with the found channel's dataframe.

### Features 
* Clicking on the "Search channels" button, the program searches for a channel according to the data entered in the "Router1" and "Router2" fields.
* When the channel is found, a notification of success will appear and the question: "Do you want to send the mail?"
* If confirmed, the program will prompt you to enter the time of the accident, the trouble ticket and the region code in the fields. A "Send the mail" button will appear.
* Сlicking on the "Send the mail" button launches Outlook and generates a letter template.
* Clicking the "Exit Program" button, close the program window.

Dependencies: pandas.read_excel, win32com.client, tkinter, re.findall.

Convert to exe: use pyinstaller or other package for create execution file