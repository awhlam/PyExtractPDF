import PyPDF2
import glob
import pandas as pd
import re

pageNum = 1
output = []
inputPDFList = glob.glob("./InputPDF/*.pdf")

# Loop through PDF files
for inputPDF in inputPDFList:
    # Open PDF file
    pdfFile = PyPDF2.PdfFileReader(open(inputPDF, 'rb'))
    # Loop through pages
    while (pageNum <= pdfFile.numPages):
        
        # Get page of current page number
        pdfPage = pdfFile.getPage(pageNum - 1)
        
        # Extract text from current page
        pdfText = pdfPage.extractText()
        
        print(pdfText)
        
        # Find ACCT_ID in text
        acct_id_match = re.search(r"\d{10}-\d", pdfText)
        if acct_id_match is not None:
            acct_id = acct_id_match.group(0)
        
        # Output to list
        output.append([inputPDF, pageNum, acct_id, pdfText])
        pageNum += 1

# Create dataframe and output to Excel File
outputDF = pd.DataFrame(output, columns = ['Filename', 'PageNum', 'Acct_ID', 'Text'])

writer = pd.ExcelWriter('output.xlsx')
outputDF.to_excel(writer,'Sheet1')
writer.save()