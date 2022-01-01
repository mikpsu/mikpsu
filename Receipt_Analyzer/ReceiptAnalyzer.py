#   _      _____ ____   _____ 
#  | |    |_   _|  _ \ / ____|
#  | |      | | | |_) | (___  
#  | |      | | |  _ < \___ \ 
#  | |____ _| |_| |_) |____) |
#  |______|_____|____/|_____/ 
# bring in libraries
import re, os, xlsxwriter

# add image processing capabilities
from types import MethodDescriptorType
from PIL import Image, ImageEnhance
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

#   __  __          _____ _   _ 
#  |  \/  |   /\   |_   _| \ | |
#  | \  / |  /  \    | | |  \| |
#  | |\/| | / /\ \   | | | . ` |
#  | |  | |/ ____ \ _| |_| |\  |
#  |_|  |_/_/    \_\_____|_| \_|

# assign directory that contains all receipt pics
directory = r'C:\Users\Mike\Documents\GitHub\mikpsu\Receipt_Analyzer\Receipts'

# create excel sheet to output information of each receipt
workbook = xlsxwriter.Workbook(r'C:\Users\Mike\Documents\GitHub\mikpsu\Receipt_Analyzer\Receipt_Information.xlsx')
worksheet = workbook.add_worksheet() # add sheet
worksheet.write('A1', 'Image File Name')
worksheet.write('B1', 'Vendor')
worksheet.write('C1', 'Date')
worksheet.write('D1', 'Payment Method')
worksheet.write('E1', 'Amount')

# define spreadsheet formats
currency_format = workbook.add_format({'num_format': '$#,##0.00'})
date_format = workbook.add_format({'num_format': 'mm/dd/yy;@'})

# iterate through each *.JPG and *.PNG file in the directory
row = 1 # index row
col = 0 # index column
for entry in os.scandir(directory):
    if (entry.path.endswith(".jpg") or entry.path.endswith(".png")) and entry.is_file():
        image_name = os.path.basename(entry.path)
        # print(image_name)
        worksheet.write(row, 0, image_name)
        receipt = Image.open(entry.path)
        
        #edit the image to increase contrast and reduce exposure
        colorEnhancer = ImageEnhance.Color(receipt)
        receipt_edit = colorEnhancer.enhance(0.0)

        brightnessEnhancer = ImageEnhance.Brightness(receipt_edit)
        receipt_edit = brightnessEnhancer.enhance(1.2)

        sharpnessEnhancer = ImageEnhance.Sharpness(receipt_edit)
        receipt_edit = sharpnessEnhancer.enhance(2.0)

        contrastEnhancer = ImageEnhance.Contrast(receipt_edit)
        receipt_edit = contrastEnhancer.enhance(4.0)

        #receipt_edit.show()

        # covert the text in the image to a string
        receipt_string = pytesseract.image_to_string(receipt_edit)

        # export the raw text to a file
        # with open(r'C:\Users\Mike\Documents\GitHub\mikpsu\Receipt_Analyzer\test_result_4.txt', mode = 'w') as file:
            #  file.write(receipt_string)

        # find the vendor of the receipt
        receipt_vendor_wild = re.compile("GIANT")
        vendor_match = receipt_vendor_wild.search(receipt_string)

        if vendor_match:
            vendor = 'Giant'
        else:
            vendor = 'Other'
        #print('Vendor:',vendor)
        worksheet.write(row, 1, vendor)

        if vendor == 'Giant':
            # find the date of the receipt
            receipt_date_wild = re.compile("\d\d/\d\d/\d\d")
            date_search = re.search(receipt_date_wild, receipt_string)

            receipt_date = date_search.group()
            receipt_dateLocation = date_search.span()
            #print('Date of receipt:', receipt_date)
            worksheet.write(row, 2, receipt_date, date_format)

            # find the payment method used
            receipt_pay_wild = re.compile("Payment.+\n")
            pay_search = re.search(receipt_pay_wild, receipt_string)
            receipt_pay = pay_search.group()
            receipt_payLocation = pay_search.span()
            card_reg_master = re.compile(".+MAS.+")

            if bool(re.match(card_reg_master, receipt_pay)):
                paymentType = "Family"
            else:
                paymentType = "Personal"

            #print('Payment Method:', paymentType)
            worksheet.write(row, 3, paymentType)

            # find the total expenditure
            receipt_total_wild = re.compile("BALA.+\d\d\.\d\d")
            total_search = re.search(receipt_total_wild, receipt_string)

            receipt_total_balance = total_search.group()
            receipt_totalLocation = total_search.span()

            receipt_totalNum_wild = re.compile("...\...")
            totalNum_search = re.search(receipt_totalNum_wild, receipt_total_balance)

            receipt_totalNum = totalNum_search.group()
            #print('Total of receipt: $', receipt_totalNum)
            worksheet.write(row, 4, float(receipt_totalNum), currency_format)

        else:
            # find the date of the receipt
            receipt_date_wild = re.compile("\d\d/\d\d/\d\d")
            date_search = re.search(receipt_date_wild, receipt_string)

            receipt_date = date_search.group()
            receipt_dateLocation = date_search.span()
            print('Date of receipt:', receipt_date)

            # find the payment method used
            receipt_pay_wild1 = re.compile("M/C")
            receipt_pay_wild2 = re.compile("MASTERCARD")
            
            pay_search = re.search(receipt_pay_wild1, receipt_string)

            # if not pay_search
            #     pay_search = re.search(receipt_pay_wild2, receipt_string)
            
            
            receipt_pay = pay_search.group()
            receipt_payLocation = pay_search.span()
            card_reg_master = re.compile(".+MAS.+")

            if bool(re.match(card_reg_master, receipt_pay)):
                paymentType = "Family"
            else:
                paymentType = "Personal"

            print('Payment Method:', paymentType)

        row = row + 1
        #end iteration
workbook.close()
# end ReceiptAnalyzer.py