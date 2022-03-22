#   _      _____ ____   _____ 
#  | |    |_   _|  _ \ / ____|
#  | |      | | | |_) | (___  
#  | |      | | |  _ < \___ \ 
#  | |____ _| |_| |_) |____) |
#  |______|_____|____/|_____/ 
# bring in libraries
import re, os, xlsxwriter, PyPDF2

# Define account
account = 'Growth'
# file path of folder containing statements
statement_path = r"C:\Users\Mike\OneDrive - The Pennsylvania State University\Finances\2021 Taxes\2021 Accounting\PNC Statements"

# Setup excel sheet
# create excel sheet to output information of each receipt
workbook = xlsxwriter.Workbook(statement_path+r'\GrowthAccountSummary.xlsx')
worksheet = workbook.add_worksheet(account) # add sheet
worksheet.write('A1', 'Period Start')
worksheet.write('B1', 'Period End')
worksheet.write('C1', 'Beginning Balance')
worksheet.write('D1', 'Deposits and Other Additions')
worksheet.write('E1', 'Checks and Other Deductions')
worksheet.write('F1', 'Ending Balance')

# define spreadsheet formats
currency_format = workbook.add_format({'num_format': '$#,##0.00'})
date_format = workbook.add_format({'num_format': 'mm/dd/yy;@'})

# Loop through each sheet
row = 1 # index row
col = 0 # index column
for entry in os.scandir(statement_path+r"\Growth"):
    # creating a pdf file object
    pdfFileObj = open(entry.path, 'rb')
    
    # creating a pdf reader object
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    page_count = pdfReader.numPages

    # creating a page object
    pageObj = pdfReader.getPage(0)
    # extracting text from page
    statement_text = pageObj.extractText()
    if 1 < page_count:
        pageObj = pdfReader.getPage(1)
        statement_text += pageObj.extractText()

    # print(statement_text)
    # find the total expenditure string
    BalanceSummary_wild = re.compile("Endingbalance.+Average monthlybalance")
    BalanceSummary_search = re.search(BalanceSummary_wild, statement_text)
    BalanceSummary_figs = BalanceSummary_search.group()
    print(BalanceSummary_figs)
    # Extract the balances
    BalanceSummary_figs = BalanceSummary_figs.replace("Endingbalance","")
    BalanceSummary_figs = BalanceSummary_figs.replace("Average monthlybalance","")
    BalanceSummary_figs_list = re.findall('\d{0,3}\,?\d{0,3}\.\d\d', BalanceSummary_figs)
    print(BalanceSummary_figs_list)
    # Extract the date
    Period_wild = re.compile("For the period.{22}")
    Period_search = re.search(Period_wild, statement_text)
    Statement_period = Period_search.group()
    
    # Extract the dates
    Statement_period_list = re.findall('\d\d\/\d\d\/\d{4}', Statement_period)
    
    # Print values to excel sheet
    worksheet.write(row, 0, Statement_period_list[0], date_format)
    worksheet.write(row, 1, Statement_period_list[1], date_format)
    for amount in range(len(BalanceSummary_figs_list)):
        num = BalanceSummary_figs_list[amount].replace(",","") 
        worksheet.write(row, (amount+2), float(num), currency_format)
    # closing the pdf file object
    pdfFileObj.close()
    row += 1
    #end iteration
workbook.close()

