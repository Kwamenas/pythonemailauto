import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import os 
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import ImageGrab
import win32com.client as win32
import pythoncom
import time
#import pandas as pd
#import dataframe_image as dfi


load_dotenv() # Load enviroment variables from .env file

#email configuration
sender_email=os.getenv("EMAIL_SENDER")
receiver_email=os.getenv("EMAIL_RECEIVER")
password=os.getenv("EMAIL_PASSWORD")

excel_file="Moda Bertha Golden Painting.xlsx"
sheet_name="Sales Receipt"

#reads the excel file 
#df=pd.read_excel("Moda Bertha Golden Painting.xlsx",sheet_name="Sales Receipt",usecols="A:H",nrows=20)

output_dir="screenshots" #creating a name for our file
os.makedirs(output_dir,exist_ok=True) #creating a dir if its not available                                                                      
image_filename="receipts.png"  #name we want to give to our image
image_path=os.path.join(output_dir,image_filename) #we join the path to create the locationfor our image

def capture_excel_range(file_path,sheet_name,image_path):
    try:
        #Initialize COM
        pythoncom.CoInitialize()

        #open excel
        excel=win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible=True #makes excel visible
        wb=excel.Workbooks.Open(os.path.abspath(file_path))
        ws=wb.Sheets(sheet_name)

        ws.Range("A1:H20").Select()

        excel.Selection.CopyPicture(Appearance=1,Format=2)

        if ImageGrab.grabclipboard() is not None:
            img=ImageGrab.grabclipboard()
            img.save(image_path)

        wb.close(False)
        excel.Quit()                                                                                                                                
    except Exception as e:
        print(f"Error capturing Excel range:{e}")
    finally:
        pythoncom.CoUninitialize()

capture_excel_range(excel_file,sheet_name,image_path)
                                             
body="Hello, please find attached your receipt thanks for your purchase"
message=MIMEMultipart()
message['From']=sender_email
message['To']=receiver_email
message['Subject']="Receipts for Sale"
message.attach(MIMEText(body,'plain'))


#Now let's attach the image
with open(image_path,"rb")as f: #open the image_path
    img=MIMEImage(f.read(),name=image_filename) #go in and read the image name in binary
    message.attach(img) #attach it to the email

try:
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender_email, password)
        server.send_message(message)
        server.send_message(message)
    print("Email sent successfuly")
except Exception as e:
    print(f"Error sending email: {e}")