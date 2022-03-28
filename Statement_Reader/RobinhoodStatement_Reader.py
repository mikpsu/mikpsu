#   _      _____ ____   _____ 
#  | |    |_   _|  _ \ / ____|
#  | |      | | | |_) | (___  
#  | |      | | |  _ < \___ \ 
#  | |____ _| |_| |_) |____) |
#  |______|_____|____/|_____/ 
# bring in libraries
import re, os, xlsxwriter, PyPDF2, slate3k as slate

# file path of folder containing statements
statement_path = r"C:\Users\Mike\OneDrive - The Pennsylvania State University\Finances\2021 Taxes\2021 Accounting\Robinhood Statements"

# Setup excel sheet
# create excel sheet to output information of each receipt
workbook = xlsxwriter.Workbook(statement_path+r'\RobinhoodAccountSummary.xlsx')
worksheet = workbook.add_worksheet() # add sheet
worksheet.write('A1', 'Period Start')
worksheet.write('B1', 'Period End')
worksheet.write('C1', 'Opening Balance')
worksheet.write('D1', 'Closing Balance')


# define spreadsheet formats
currency_format = workbook.add_format({'num_format': '$#,##0.00'})
date_format = workbook.add_format({'num_format': 'mm/dd/yy;@'})

# Loop through each sheet
row = 1 # index row
col = 0 # index column
for entry in os.scandir(statement_path):
    # creating a pdf file object
    pdfFileObj = open(entry.path, 'rb')
    doc = slate.PDF(pdfFileObj)

    # extracting text from page
    statement_text = doc[0]
    statement_text = statement_text.replace("$","")
    statement_text = statement_text.replace(" ","")
    statement_text = statement_text.replace("\n\n","\n")
    statement_text = statement_text.replace(",","")
    # print(statement_text)
    # find the total expenditure string
    BalanceSummary_wild = re.compile("\d.+\n\d.+\n\d.+\n\d.+\n\d.+\n\d.+")
    BalanceSummary_search = re.search(BalanceSummary_wild, statement_text)
    BalanceSummary_figs = BalanceSummary_search.group()
    # print(BalanceSummary_figs)
    # Extract the balances
    BalanceSummary_figs_list = BalanceSummary_figs.split("\n")
    del BalanceSummary_figs_list[:4]
    # print(BalanceSummary_figs_list)
    # Extract the date
    Period_wild = re.compile("\d\d\/\d\d\/\d{4}to\d\d\/\d\d\/\d{4}")
    Period_search = re.search(Period_wild, statement_text)
    Statement_period = Period_search.group()
    # print(Statement_period)
    # Extract the dates
    Statement_period_list = re.findall('\d\d\/\d\d\/\d{4}', Statement_period)
    print(Statement_period_list)
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

