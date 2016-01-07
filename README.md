# MailToWI
 Useful application for the automatic management of reports received via e-mail that are transformed on Bugs
 with the original e-mail text on attachment.
 
 The application is implemented in C # and .NET 4.5, and must be performed as a process on a Windows Server (possibly different from MS-TFS). To activate it as a process, we recommend using the sc command with the create option (see https://technet.microsoft.com/it-it/library/cc990289(v=ws.10).aspx)

In order to use the application you must:
• Have a service user on the domain with a dedicated mailbox to receive email from turning into bugs
• Provide the user with write access to the project MS-TFS
Whenever the user receives an e-mail the service will pick up and turns it into bug on MS-TFS.

You can place an alert on TFS  that  send to your mailbox service an email when the WorkItem passes in a particular state (eg. Done). 
In this case, the service extracts the name of the user who sent the message from the text of the email attachment of the work item and 
sends the information to the original sender.
