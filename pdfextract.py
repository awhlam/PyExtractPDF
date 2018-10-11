import PyPDF2
import glob
import pandas as pd
import re

output = []
inputPDFList = glob.glob(".\input\*.pdf")

# Loop through PDF files
for inputPDF in inputPDFList:
    # Open PDF file
    pdfFile = PyPDF2.PdfFileReader(open(inputPDF, 'rb'))
    # Loop through pages
    pageNum = 1
    while (pageNum <= pdfFile.numPages):
        
        # Get page of current page number
        pdfPage = pdfFile.getPage(pageNum - 1)
        
        # Extract text from current page
        pdfText = pdfPage.extractText()
        
        # Find ACCT_ID in text
        acct_id_match = re.search(r"\d{10}-\d", pdfText)
        if acct_id_match is not None:
            acct_id = acct_id_match.group(0)[:10]
        else:
            acct_id = ''
            
        # Find Statement Date
        statement_date_match = re.search(r"\d{2}\/\d{2}\/\d{4} \d{2}\/\d{2}\/\d{4}", pdfText)
        if statement_date_match is not None:
            statement_date = statement_date_match.group(0)[:10]
        else:
            statement_date_match = re.search(r"Statement Date[:|;] \d{2}\/\d{2}\/\d{4}", pdfText)
            if statement_date_match is not None:
                statement_date = statement_date_match.group(0)[16:27]
            else:
                statement_date = ''
        
        # Output to list
        output.append([inputPDF, pageNum, acct_id, statement_date, pdfText])
        pageNum += 1

# Create dataframe and output to Excel File
outputDF = pd.DataFrame(output, columns = ['FILENAME', 'PAGE_NUMBER', 'ACCT_ID', 'STATEMENT_DATE', 'TEXT'])
writer = pd.ExcelWriter('./output/output.xlsx')
outputDF.to_excel(writer,'Sheet1')
writer.save()