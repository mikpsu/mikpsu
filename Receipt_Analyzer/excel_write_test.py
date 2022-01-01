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
worksheet.write('A2', 'Vendor')
worksheet.write('A3', 'Date')
worksheet.write('A4', 'Amount')
worksheet.write('A5', 'Payment Method')
workbook.close()