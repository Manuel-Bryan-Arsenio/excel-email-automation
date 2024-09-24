import pandas as pd # to extract datas from excel
import smtplib  # to send email
import numpy as np # to replace NaN
import sys  # to exit the code once an error is found

data = pd.read_excel("~/Documents/CompSci_Independent_Study/GitHub-Uploads/excel-email-automation-py/Excel_Sample_Data.xlsx") # reading the datas in excel

data.replace({"Start": {np.nan: 'TBD'}}, inplace=True)  # for the column data, NaN will be replace with a string "TBD"
data.replace({"Online/Offline": {np.nan: 'none'}}, inplace=True)  # for the column data, NaN will be replace with a string "none"
    # line 8 will be crucial for the if, elif and else condition in the sending email code, as .lower() only works with strings not NaN

class Candidate:    # candidate class
    
    def __init__(self, date, time, form, first, last, email, location, link):   # attributes of the candidate class
        self.form = form
        self.first = first
        self.last = last
        self.email = email
        self.location = location
        self.date = date
        self.time = time
        self.link = link
    
    def formfullname(self): # function for "Mr. Firstname Lastname" or "Ms. Firstname Lastname"
        if pd.isna(self.first) == True: 
            return f"{self.form} {self.last}"
        elif pd.isna(self.last) == True:
            return f"{self.form} {self.first}"
        else:
            return f"{self.form} {self.first} {self.last}"

data_list = []  # Empty list to store the data from excel in a list

for i in range(len(data)):  # for loop iterating over every row in the table
    data_list.append(
        Candidate(
            data.loc[i, data.columns[0]],    
            data.loc[i, data.columns[1]],              
            data.loc[i, data.columns[2]],               
            data.loc[i, data.columns[3]],              
            data.loc[i, data.columns[4]],     
            data.loc[i, data.columns[5]],               
            data.loc[i, data.columns[6]],              
            data.loc[i, data.columns[7]]        
        )
    )

# Checking for similar time and date between candidates
for i in range(len(data_list)): # range(len(data_list)) = (0, len(data_List))
    for j in (i+1, len(data_list)-1):
        if data_list[i].date == data_list[j].date and data_list[i].time == data_list[j].time:
            sys.exit(f"ERROR. {data_list[i].formfullname()} and {data_list[j].formfullname()} has the same interview time and date.")  # Exit the program and print the text
        else:
            continue    # Continue the code if no same interview time and date between candidates
    break   # This breaks the outer loop as well once a conflict is found    

# Checking for candidates location
for data in data_list:
    if data.location.lower() == "none": #   if location cell of the candidate is empty, NaN which is replace with the string "none", line 9.
        sys.exit(f"no location found for {data.formfullname()}")    # Exit the program and print error message
    else:
        continue

# Sending Email
for data in data_list:  # for loop iterating over every instances in the list.

    sender_email = "sender@gmail.com"   # sender's email
    gmail_password = "my_gmail_app_password_123"    # sender's outlook app password

    if data.location.lower() == "online":   # Email for online interview
        sender_email  # sender's Email
        receiver_email = data.email # candidate's Email

        # Subject and Message
        subject = "Invitation to the aptitude assessment interview at University A, winter term 2024/25"
        message_1 = f"Dear {data.formfullname()},"
        message_2 = "Thank you for applying for the Bachelors program Aerospace at the University A."
        message_3 = "We are pleased to invite you to an interview."
        message_4 = "The winter term 2024/25 aptitude assessment interviews will be held digitally via a video conferencing system."
        message_5 = "If you don't agree to interview in a digital form, please inform us by email. We will then offer you a personal interview on campus later."
        message_6 = "Information about the interview via video conference:"
        message_7 = f"Your interview will take place on {data.date} at {data.time}. and will be conducted using the B video conferencing service."
        message_8 = f"Please follow this link on the date mentioned above to join the meeting:\n\n{data.link}"
        message_9 = "If you disagree with the digital format, please inform us by email."
        message_10 = "For the implementation, it is necessary that you are connected via video during the entire conversation period."
        message_11 = "Therefore, ensure a terminal with a functioning video and microphone function and a stable internet connection at your appointment."
        message_12 = "Please have your ID card or passport ready for identity verification at the beginning of the conversation."
        message_13 = "The interview will be in English.\n\nPlease confirm the receipt of this email with a short reply.\n\nBest regards"

        # Email's content
        text = f"Subject: {subject}\n\n{message_1}\n\n{message_2}\n{message_3}\n\n{message_4}\n{message_5}\n\n{message_6}\n\n{message_7}\n\n{message_8}\n\n{message_9}\n{message_10}\n{message_11}\n{message_12}\n\n{message_13}"

        server = smtplib.SMTP("smtp.gmail.com", 587) # Port 587 is the standard port used for SMTP (Simple Mail Transfer Protocol) submission with TLS encryption.
        server.starttls()   # .starttls refer to the use of TLS encryption

        server.login(sender_email, gmail_password)    # login to sender's outlook server
        
        server.sendmail(sender_email, receiver_email, text) # executing the sending of the email

        print(f"Email has been sent to {receiver_email}")   # to print and indicate that the Email has been sent.

    elif data.location.lower() == "offline": # Email for offline interview
        sender_email    # sender's Email
        receiver_email = data.email # candidate's Email

        # Subject and Message
        subject = "Invitation to the aptitude assessment interview at University A, winter term 2024/25"
        message_1 = f"Dear {data.formfullname()},"
        message_2 = "Thank you for applying for the Bachelors program Aerospace at the University A."
        message_3 = "We are pleased to invite you to an interview."
        message_4 = "The winter term 2024/25 aptitude assessment interviews will be held offline, as requested."
        message_5 = "If you don't agree to interview in the offline form, please inform us by email. We will then offer you a personal interview digitally later."
        message_6 = "Information about the interview:"
        message_7 = f"Your interview will take place on {data.date} at {data.time}. and will be conducted at Location C."
        message_8 = "If you disagree with the offline format, please inform us by email."
        message_9 = "The waiting room at the location opens 15 minutes before the interview."
        message_10 = "Please have your ID card or passport ready for identity verification at the beginning of the conversation."
        message_11 = "The interview will be in English."
        message_12 = "TPlease confirm the receipt of this email with a short reply.\n\nBest regards"

        # Email's content
        text = f"Subject: {subject}\n\n{message_1}\n\n{message_2}\n{message_3}\n\n{message_4}\n{message_5}\n\n{message_6}\n\n{message_7}\n\n{message_8}\n{message_9}\n{message_10}\n{message_11}\n\n{message_12}"

        server = smtplib.SMTP("smtp.gmail.com", 587) # Port 587 is the standard port used for SMTP (Simple Mail Transfer Protocol) submission with TLS encryption.
        server.starttls()   # .starttls refer to the use of TLS encryption

        server.login(sender_email, gmail_password)    # login to sender's outlook server
        
        server.sendmail(sender_email, receiver_email, text) # executing the sending of the email

        print(f"Email has been sent to {receiver_email}")   # to print and indicate that the Email has been sent.

# NOTES for further development
#
# Line 48 - 55: the loop only checks between 2 candidates, so if there are more than 2 candidates with the same interview time and date we need to run the program again and change it again.
#
# Line 57 - 127: too much message variable, unnecessarily complex.
#   [Solution] Message/text can be written separately and can be imported to be used in this code.