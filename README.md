# Simple Gmail Sender
Simple tool for spam sending with attachment through xls list of recepients

# Python 
Python 3.7

# Usage 
Fill ./files/my_list_of_recepients.xls with recepients, set flag 'is_needed' to '1' for actual recepients \
Replace ./files/my_attachment.pdf with actual attachment

To start sending use
```>>>GMailer().send_mails()```

To import mail chain to Word file 'some_recepient@somemail.com_01-01-2020_15:00.docx' for each recepient who replied use 
```>>>GMailer().update_chains()```
