# Inspired by https://towardsdatascience.com/building-a-simple-text-recognizer-in-python-93e453ddb759

# bring in regular expression library
import re

# add image processing capabilities
from types import MethodDescriptorType
from PIL import Image

# will convert the image to test string
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# assign image from the source path
receipt = Image.open(r'C:\Users\Mike\Documents\GitHub\mikpsu\Receipt_Analyzer\Giant_Receipt_edited.jpg')

# convert the image to result and save result into variable
receipt_string = pytesseract.image_to_string(receipt)
#print('Raw output:**********************\n', result, '**********************')

# convert string to list of strings with each word
result_list = receipt_string.split()
print(result_list)

# find the element 
# export results into text file
with open(r'C:\Users\Mike\Documents\GitHub\mikpsu\Receipt_Analyzer\test_result.txt', mode = 'w') as file:
    file.write(receipt_string)

# extract desired information:
    # Vendor

    # find date using regex
    receipt_date = re.compile(".+/.+/.+")

    date_location = receipt_string.find(re.match(receipt_date, receipt_string))
    print(date_location)

    # Payment Method
    # Total Expense
