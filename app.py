# Insert your python code here
# IMPORT PACKAGES IN THIS CELL

import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

#Question 1
# INSERT ANSWER FOR QUESTION 1 HERE

while True:
    directory = single_input('Please enter the path for the directory where the files are located')
    directory = directory.strip('"')
    if directory[-1] != '\\':
        directory = directory + '\\'
    try:
        filenames = os.listdir(directory)
        if ('ProductData.xlsx' in filenames) and ('MasterInvoice.pdf' in filenames) and ('Order1.xlsx' in filenames) and ('Order2.xlsx' in filenames) and ('Order3.xlsx' in filenames) and ('Order4.xlsx' in filenames):
            break
        else:
            show_message('Error', 'The directory that you entered is invalid. Please try again.')
    except:
            show_message('Error', 'The directory that you entered is invalid. Please try again.')


# Question 2
# INSERT ANSWER FOR QUESTION 2 HERE
productdf = pd.read_excel(directory + 'ProductData.xlsx',sheet_name='Descriptions')
product_descriptions = {}
for index, row in productdf.iterrows():
    dvalue = row.Description
    dkey = row['Unit number']
    product_descriptions[dkey] = dvalue

pricesdf = pd.read_excel(directory + 'ProductData.xlsx',sheet_name='Prices')
product_prices = {}
for index, row in pricesdf.iterrows():
    pvalue = row['Price per unit']
    pkey = row['Unit number']
    product_prices[pkey] = pvalue

#Question 3
# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num,filename,directory):
    invoice_num = int(invoice_num) #makes sure that the invoice num is an integer

    order_data = {}

    invoice_path = directory + filename
    
    template = openpyxl.load_workbook(invoice_path) #getting the name, email, and date from the order form
    selected_sheet = template['OrderForm']
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    wrongdate = selected_sheet['C9'].value
    month = wrongdate.month
    day = wrongdate.day
    year = wrongdate.year
    date = f'{month}/{day}/{year}'
    street = selected_sheet['C16'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zipcode = selected_sheet['C19'].value
    address1 = f'{street} \n{city}, {state} {zipcode}'
    company = selected_sheet['C15'].value

    

    
    product_key1 = selected_sheet['B24'].value
    quantity1 = selected_sheet['C24'].value
    price1 = product_prices[product_key1]
    price1str = f'${price1: .2f}'
    amount1 = quantity1 * price1
    amount1str = f'${amount1: .2f}'
    description1 = product_descriptions[product_key1]
    total = amount1
    
    if selected_sheet['B25'].value != None:
        product_key2 = selected_sheet['B25'].value
        quantity2 = selected_sheet['C25'].value
        price2 = product_prices[product_key2]
        price2str = f'${price2: .2f}'
        amount2 = price2 * quantity2
        amount2str = f'${amount2: .2f}'
        description2 = product_descriptions[product_key2]
        total = total + amount2
    else:
        product_key2 = ''
        quantity2 = ''
        price2str = ''
        amount2str = ''
        description2 = ''

    if selected_sheet['B26'].value != None:
        product_key3 = selected_sheet['B26'].value
        quantity3 = selected_sheet['C26'].value
        price3 = product_prices[product_key3]
        price3str = f'${price3: .2f}'
        amount3 = price3 * quantity3
        amount3str = f'${amount3: .2f}'
        description3 = product_descriptions[product_key3]
        total = total + amount3
    else:
        product_key3 = ''
        quantity3 = ''
        amount3str = ''
        price3str = ''
        description3 = ''
    
    totalstr = f'${total: .2f}'

    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['Total'] = totalstr
    order_data['CustomerAddress1'] = address1
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = company
    order_data['CustomerAddress2'] = ''
    order_data['Price0'] = price1str
    order_data['Quantity0'] = str(quantity1)
    order_data['Amount0'] = amount1str
    order_data['Description0'] = description1
    order_data['Price1'] = price2str
    order_data['Quantity1'] = str(quantity2)
    order_data['Amount1'] = amount2str
    order_data['Description1'] = description2
    order_data['Price2'] = price3str
    order_data['Quantity2'] = str(quantity3)
    order_data['Amount2'] = amount3str
    order_data['Description2'] = description3


    return invoice_num, name, email, date, order_data

#Question 4
# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num,order_data,directory):
    create_pdf('MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)

#Question 5
# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):

    # Enter your gmail user name and password

    gmail_user = "eggetsto000@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()

    # List of recipients

    recipients = ["eggetsto000@gmail.com"]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    name = 'Stockton'

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    email: sales@pythontechco.com"""
    #triple quotes allow a string to span multiple lines without using a new line character
    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    #path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

#Question 6
# INSERT ANSWER FOR QUESTION 6 HERE
invoice_num = 8515
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        invoice_num, name, email, date, order_data = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1