I wrote this VB program to advise me if any new messages are received in to my Outlook Express.

It advises me by sending an SMS test message to my mobile phone.

The message looks like:

"A new message from "John Citizen" with the subject "Important information" has been received."

Basically this program downloads each message, strips the Sender's friendly name and the Subject header from the incoming mail message and, using the MobileFBUS.OCX file included, sends a text message to the specified mobile phone number. I think MobileFBUS.OCX only works with Nokia mobile phones. Obviously, the receiving mobile phone can be any brand you like.

Prerequisites:

Outlook Express

A Nokia digital mobile phone with a data cable connected to a COM port on the PC that will run the program.

The installation of the "Long Timer" OCX and the MobileFBUS OCX.


The .INI file contains the variables.

MobilePort=COM1

This is the port the Nokia phone's data cable is connected to.

CheckInterval=10

This is specified in minutes. Internally, the program multiplies any number given here by 60,000 (60,000 = one minute). Therefore, putting in '10' here would cause the program to wait 10 minutes before each check. The minimum is 1.

MobileNumber=1234567890

You put the number of the mobile you want to receive the messages here.

AccountName=yourmailaccount

This is the account name you use to 'log on' to your Outlook Express.

AccountPassword=yourmailpassword

...and, obviously, this is your Outlook Express password.



Don't forget that Outlook Express is simply a 'front-end' to the MAPI subsystem. Therefore, any messages received while this program is running will subsequently appear in an unread state in Outlook Express when next you run OE.

Enjoy!

Mark

marximus27@hotmail.com
