# Outlook-SendDuringWorkHours

Use this macro to delay messages to the next working day.  I used this daily to ensure meesages sent after hours in Outlook automatically set the delayed send time.  

To enable this macro follow these sets. 

From Microsoft Outlook, open the VBA Editor. Place code in ThisOutlookSession.  Constant values can be set to your desired work hours for your time zone. 

c_WorkHourStart -- This is the earliest you will send an email.  Emails after work hours will be delayed until this time

c_WorkHourEnd   -- This is the latest you will send an email.  Emails after work hours will be delayed

c_BypassForHighImportance -- If a message is set to High importance, the message is not delayed and the code exits.
