# PyGraphInbox

Get EXO mail messages via Graph using Python.  
Messages are pulled from the inbox of the users.  

# Files

| fileName | Description |
|----------|-------------|
|Inbox.py  | Main script file. Connect to each mailbox listed in the 'users.txt' file and output the messages to the console. |  
|Requirements.txt | Requirements file listing required modules to be installed in order to run the script. |  
| users.txt | Txt file with the list of primary smtp addresses of the mailboxes the script will work on. |
| config.cfg | Config file. Required lines are:<br>**ClientID**<br>**TenantID**<br>**ClientSecret**<br>**authtenant**<br>**userslistfilename**<br> |
- _Note that the last line of the config.cfg file, references the text file with the list of email addresses. You can modify this line to map to a different file name._  
- _You need to have an Azure App registered. If you want to know how to create the app, you can follow these [steps.](https://docs.microsoft.com/en-us/graph/tutorials/python?tabs=aad&tutorial-step=7)_  

# How to run

In my testing I used Windows.  
You can open Command Prompt or Powershell and run:  
```powershell
Python.exe inbox.py  
```
it will read the config file, and lookup for the userslistfilename file to fetch the list of mailboxes.  
<br>
You can pass the email addresses as arguments:  
```powershell
Python.exe inbox.py "user1@contoso.com"
```
or even:  
```powershell
Python.exe inbox.py "user1@contoso.com","user2@contoso.com"
```
> Passing addresses as arguments, will take precedence even if the line 'userslistfilename' in the config file exists.