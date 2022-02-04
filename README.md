# Email-Sending-using-Python
Sending mail from Outlook using Python (Pywin32)

We are sending mail from outlook using python. 
We are using Pywin32 module which is use to interact with Windows API..

We are using pywin32 to send mail from outlook.

Imports and Dispatch
The first step is to import the required library and dispatch an instance of Outlook.

import win32com.client as client

outlook = client.Dispatch('Outlook.Application')
Create a Mail Item
Now that you have an Outlook instance, you can create a mail item. Other than mail items, there are many other items that you can create. You can find the enumerations on Microsoft's website.

While it's not necessary, we're going to use the Display method to open the message window so that we can see the changes in Outlook as we're making them.

message = outlook.CreateItem(0) # 0 is the code for a mail item (see the enumerations)
message.Display()
Adjusting message properties
After creating a mail item, you are ready to start setting its properties.

The message recipients are set with the To, CC, and BCC properties.

The Subject is self explanatory.

There are two options for the main text of the email:

Body is the property for a plain-text formatted email body
HTMLBody is the property for an html formatted email body
message.To = 'bob.mortimer@email.com'
message.CC = 'james.acaster@email.com'
message.BCC = 'richard.ayoade@email.com'

message.Subject = 'Happy Birthday"'
message.Body = 'Wishing you a very happy birthday!!'
Other settings
You can also set the SentOnBehalfOfName property if you have permission to send on behalf of another party.

message.SentOnBehalfOfName = 'david.mitchell@email.com'
Sending and Saving
When you are ready to save or send, you can use the Save method to save the message to the drafts folder. The Send method will send the current mail item.

By default, an email will be sent from the email address of the profile that is currently logged on.

message.Save() # save to drafts folder
message.Send() # send to outbox
