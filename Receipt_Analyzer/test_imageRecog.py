# bring in regular expression library
import re, os

# add image processing capabilities
from types import MethodDescriptorType
from PIL import Image, ImageEnhance

# convert the image to test string
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# assign image from the source path
receipt = Image.open(r'C:\Users\Mike\Documents\GitHub\mikpsu\Receipt_Analyzer\Giant_Receipt_3.jpg')
receipt.show()
#edit the image to increase contrast and reduce exposure
colorEnhancer = ImageEnhance.Color(receipt)
receipt_edit = colorEnhancer.enhance(0.0) # make black and white with 0.0

brightnessEnhancer = ImageEnhance.Brightness(receipt_edit)
receipt_edit = brightnessEnhancer.enhance(2.0)

sharpnessEnhancer = ImageEnhance.Sharpness(receipt_edit)
receipt_edit = sharpnessEnhancer.enhance(2.0)

contrastEnhancer = ImageEnhance.Contrast(receipt_edit)
receipt_edit = contrastEnhancer.enhance(3.0)

receipt_edit.show()

# convert the image to result and save result into variable
receipt_string = pytesseract.image_to_string(receipt)
#print('Raw output:**********************\n', result, '**********************')

# find the element 
# export results into text file
with open(r'C:\Users\Mike\Documents\GitHub\mikpsu\Receipt_Analyzer\test_result_3.txt', mode = 'w') as file:
    file.write(receipt_string)

#*********************** extract desired information **********************

# check the vendor
receipt_vendor_wild = re.compile("GIANT")
vendor_match = receipt_vendor_wild.search(receipt_string)

if vendor_match:
    vendor = 'Giant'
else:
    vendor = 'Other'
print('Vendor:',vendor)

# find date using regex
receipt_date_wild = re.compile("../../..")

date_search = re.search(receipt_date_wild, receipt_string)

receipt_date = date_search.group()
receipt_dateLocation = date_search.span()
print('Date of receipt:', receipt_date)

# Payment Method
receipt_pay_wild = re.compile("Payment.+\n")


pay_search = re.search(receipt_pay_wild, receipt_string)

receipt_pay = pay_search.group()
# print(receipt_pay)
receipt_payLocation = pay_search.span()

card_reg_master = re.compile(".+MAS.+")

if bool(re.match(card_reg_master, receipt_pay)):
    paymentType = "Family"
else:
    paymentType = "Personal"

print('Payment Method:', paymentType)
# Total Expense
receipt_total_wild = re.compile("BALA.+\...")
total_search = re.search(receipt_total_wild, receipt_string)

receipt_total_balance = total_search.group()
receipt_totalLocation = total_search.span()

receipt_totalNum_wild = re.compile("...\...")
totalNum_search = re.search(receipt_totalNum_wild, receipt_total_balance)

receipt_totalNum = totalNum_search.group()
print('Total of receipt: $', receipt_totalNum)