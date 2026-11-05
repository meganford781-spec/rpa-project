import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
import datetime as dt
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# question 1
while True:
    directory = single_input("Insert directory where files are located", mask=False, width=400, height=200)
    directory = directory.strip('"').strip("'")
    if directory[-1] == '/':
        break
    else:
        directory = directory + '/'
    filenames = os.listdir(directory)

    filenames = os.listdir(directory)
    if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
        break
    else:
        show_message("Error",'files not included in directory')
print(directory)
print(filenames)# Insert your python code here

# question 2
filepath = directory+'ProductData.xlsx'
descriptions = pd.read_excel(filepath, sheet_name = 'Descriptions')
prices = pd.read_excel(filepath, sheet_name = 'Prices')

product_descriptions = {}
product_prices = {}

for index,row in descriptions.iterrows():
    descript = row['Description']
    unit_num = row['Unit number']
    product_descriptions[unit_num] = descript

for index,row in prices.iterrows():
    price = row['Price per unit']
    unit_num = row['Unit number']
    product_prices[unit_num] = price

print(product_descriptions)
print(product_prices)

#Question 3
df = get_pdf('MasterInvoice.pdf')
df
order_data = {}
def get_order_data(invoice_num,filename,directory):
    order_file = openpyxl.load_workbook(directory+filename, data_only=True)
    selected_sheet = order_file['OrderForm']
    order_data['InvoiceNum'] = str(invoice_num)
    date = (selected_sheet['C9'].value).strftime('%m/%d/%Y')
    order_data['Date'] = date
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value

    order_data['CustomerAddress1']=str(selected_sheet['C16'].value)
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {selected_sheet['C19'].value}'
    
    order_data['Price0'] = str(product_prices[selected_sheet['B24'].value])
    order_data['Quantity0'] = str(selected_sheet['C24'].value)
    amount0 = product_prices[selected_sheet['B24'].value] * selected_sheet['C24'].value
    order_data['Amount0'] = f'${amount0:,.2f}'
    order_data['Description0'] = product_descriptions[selected_sheet['B24'].value]
    total = amount0
    
    if pd.isnull(selected_sheet['B25'].value) == False:
        order_data['Price1'] = str(product_prices[selected_sheet['B25'].value])
        order_data['Quantity1'] = str(selected_sheet['C25'].value)
        amount1 = product_prices[selected_sheet['B25'].value] * selected_sheet['C25'].value
        order_data['Amount1'] = f'${amount1:,.2f}'
        order_data['Description1'] = product_descriptions[selected_sheet['B25'].value]
        total = total + amount1

    if pd.isnull(selected_sheet['B26'].value) == False:
        order_data['Price2'] = str(product_prices[selected_sheet['B26'].value])
        order_data['Quantity2'] = str(selected_sheet['C26'].value)
        amount2 = product_prices[selected_sheet['B26'].value] * selected_sheet['C26'].value
        order_data['Amount2'] = f'${amount2:,.2f}'
        order_data['Description2'] = product_descriptions[selected_sheet['B26'].value]
        total = total + amount2
    order_data['Total'] = f'${total:,.2f}'
    
    return order_data, name, email, date
get_order_data(8515,'Order1.xlsx',directory)

# Question 4
def create_invoice(invoice_num,order_data,directory):
    create_pdf(directory+'MasterInvoice.pdf', directory+f'Invoice {invoice_num}.pdf', order_data)

create_invoice(8515,order_data,directory)

#Question 5
def send_email(invoice_num,name,email,date,directory):
    gmail_user = "meganford781@gmail.com"
  
    # gmail_password_path = select_file('Please select the gmail password file')
    # gmail_password = open(gmail_password_path).read()
    body = f"""Dear {name},
    
Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84602
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = email
    message['Subject'] = f'Your Order from PythonTech Co. - Invoice {invoice_num}'
    message.set_content(body)

    path_to_file = directory+f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()

# Question 6
filenames = os.listdir(directory)
filenames.sort()
print(filenames)
invoice_num = 8515
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()
for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        try: 
            order_data,name,email,date = get_order_data(invoice_num,file,directory)
            create_invoice(invoice_num,order_data,directory)
            send_email(invoice_num,name,email,date,directory)
            invoice_num = invoice_num+1
        except:
            pass
