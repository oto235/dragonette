import pandas as pd
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication 
from email.mime.multipart import MIMEMultipart 
import smtplib 
import os 

# smtp = smtplib.SMTP('smtp.gmail.com', 587) 
# smtp.ehlo() 
# smtp.starttls()

# smtp.login('oto235@gmail.com', 'lgth bgnp rctu uicc')


# def message(subject="2023 Tax donation receipt Dragonettes",  
#             text="", img=None, 
#             attachment=None): 
    
#     # build message contents 
#     msg = MIMEMultipart() 
      
#     # Add Subject 
#     msg['Subject'] = subject   
      
#     # Add text contents 
#     msg.attach(MIMEText(text)) 

#     if attachment is not None:

#         if type(attachment) is not list: 
            
#                 # if it isn't a list, make it one 
#             attachment = [attachment]   

#         for one_attachment in attachment: 

#             with open(one_attachment, 'rb') as f: 
                
#                 # Read in the attachment 
#                 # using MIMEApplication 
#                 file = MIMEApplication( 
#                     f.read(), 
#                     name=os.path.basename(one_attachment) 
#                 ) 
#             file['Content-Disposition'] = f'attachment; filename="{os.path.basename(one_attachment)}"'
            
#             msg.attach(file)

#     return msg



# # Call the message function 
# msg = message(text= "Hi there!", 
#               attachment= r"C:\Users\oto23\OneDrive\Documents\School_kids\dragonette_scripts\dragonette\2023 Tax donation receipt The Perry House.pdf") 

# email = 'jasonschwarzpt@gmail.com'  
# # Make a list of emails, where you wanna send mail 
# to = [email] 
  
# # Provide some data to the sendmail function! 
# smtp.sendmail(from_addr="oto235@gmail.com", 
#               to_addrs=to, msg=msg.as_string()) 
  
#  # Finally, don't forget to close the connection 
# smtp.quit()





df = pd.read_csv("2023_Taxes_131.csv")

# print(df)

start = 5
stop = len(df)

for idx in range(start, stop):
    email = df['email'][idx]
    print(email)



