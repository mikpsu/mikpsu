#   _      _____ ____   _____ 
#  | |    |_   _|  _ \ / ____|
#  | |      | | | |_) | (___  
#  | |      | | |  _ < \___ \ 
#  | |____ _| |_| |_) |____) |
#  |______|_____|____/|_____/ 
# bring in libraries
import re, os, xlsxwriter, PyPDF2, slate3k as slate

# file path of folder containing statements
statement_path = r"C:\Users\Mike\OneDrive - The Pennsylvania State University\Finances\2021 Taxes\2021 Accounting\Visa Statements"

# Setup excel sheet
# create excel sheet to output information of each receipt
workbook = xlsxwriter.Workbook(statement_path+r'\VisaExpenses.xlsx')
worksheet = workbook.add_worksheet() # add sheet
worksheet.write('A1', 'Statement Period End')
worksheet.write('B1', 'Purchase Activity')

# define spreadsheet formats
currency_format = workbook.add_format({'num_format': '$#,##0.00'})
date_format = workbook.add_format({'num_format': 'mm/dd/yy;@'})

# Loop through each sheet
row = 1 # index row
col = 0 # index column
for entry in os.scandir(statement_path):

    # creating a pdf file object
    pdfFileObj = open(entry.path, 'rb')
    
    # Alternate method: slate
    # doc = slate.PDF(pdfFileObj)

    # # extracting text from page
    # statement_text = doc[1]

    # creating a pdf reader object
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    page_count = pdfReader.numPages

    # creating a page object
    pageObj = pdfReader.getPage(0)
    # extracting text from page
    statement_text = pageObj.extractText()

    statement_text = statement_text.replace("$","")
    statement_text = statement_text.replace(" ","")
    statement_text = statement_text.replace("\n\n","\n")
    statement_text = statement_text.replace(",","")

    # print(statement_text)
    # Looking for: "BillPeriod:02/01/2021-02/28/2021PaymentInformationNewBalance262.00"
    # find the total expenditure string
    NewBalance_wild = re.compile("NewBalance\(?\d{1,4}\.\d{2}\)?")
    BalanceSummary_search = re.search(NewBalance_wild, statement_text)
    BalanceSummary_figs = BalanceSummary_search.group()
    # Extract the dollar amount
    BalanceSummary_figs = BalanceSummary_figs.replace("NewBalance","")

    # convert parenthetical number to negaive
    if (BalanceSummary_figs.find('(') != -1):
        BalanceSummary_figs = BalanceSummary_figs.replace("(","-")
        BalanceSummary_figs = BalanceSummary_figs.replace(")","")
    
    print(BalanceSummary_figs)

    # Extract the date
    Period_wild = re.compile("BillPeriod:\d{2}\/\d{2}\/\d{4}-\d{2}\/\d{2}\/\d{4}")
    Period_search = re.search(Period_wild, statement_text)
    Statement_period = Period_search.group()

    # Extract the dates
    Statement_period = Statement_period[(11+2+1+2+1+4+1):]
    # print(Statement_period)

    # Print values to excel sheet
    worksheet.write(row, 0, Statement_period, date_format)
    worksheet.write(row, 1, float(BalanceSummary_figs), currency_format)
    # closing the pdf file object
    pdfFileObj.close()
    row += 1
    #end iteration
workbook.close()

