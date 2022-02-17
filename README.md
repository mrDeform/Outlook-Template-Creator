# Outlook template creator
A program for searching for a lease channel using Excel tables and preparing a template for sending a letter.

Important:
the program works with an Excel files containing information about the channels. The search for channels is carried out by columns with the name of the equipment at the connection points. If the search for information in the first file was unsuccessful, the program proceeds to search for information in the second file. The received information contains email addresses for communication. An Outlook template is then created for the received email address with the found channel's dataframe. The file "dict.csv" is used to get the area code, which is used to send a letter to the responsible branch.

### Features 
* Clicking on the "Search channels" button, the program searches for a channel according to the data entered in the "Router1" and "Router2" fields.
* When the channel is found, a notification of success will appear and the question: "Do you want to send the mail?"
* If confirmed, the program will offer to add information about the crash time and ticket accident. A "Send email" button will appear.
* Сlicking on the "Send the mail" button launches Outlook and generates a letter template.
* Clicking the "Restart the program" button, restarts the program window.
* Clicking the "Exit Program" button, close the program window.

Dependencies: win32com.client, pandas.read_excel(read_csv), tkinter, re.findall, os.path.isfile.

Convert to exe: use pyinstaller or other package for create execution file