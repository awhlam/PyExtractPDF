import PyPDF2
import glob
import pandas as pd
import re

def get_acct_id(pdfText):
	acct_id_match = re.search(r"\d{10}-\d", pdfText)
	if acct_id_match is not None:
		return acct_id_match.group(0)[:10]
	else:
		return ''

def get_statement_dt(pdfText):
	statement_date_match = re.search(r"\d{2}\/\d{2}\/\d{4} \d{2}\/\d{2}\/\d{4}", pdfText)
	if statement_date_match is not None:
		return statement_date_match.group(0)[:10]
	else:
		statement_date_match = re.search(r"Statement Date[:|;] \d{2}\/\d{2}\/\d{4}", pdfText)
		if statement_date_match is not None:
			return statement_date_match.group(0)[16:27]
		else:
			return ''

def write_to_file():
	# Create dataframe and output to Excel File
	outputDF = pd.DataFrame(output, columns = ['FILENAME', 'PAGE_NUMBER', 'ACCT_ID', 'STATEMENT_DATE', 'TEXT'])
	writer = pd.ExcelWriter('./output/output.xlsx')
	outputDF.to_excel(writer,'Sheet1')
	writer.save()

if __name__ == "__main__":
	output = []
	inputPDFList = glob.glob(".\input\*.pdf")

	# Loop through PDF files
	for inputPDF in inputPDFList:
		# Open PDF file
		pdfFile = PyPDF2.PdfFileReader(open(inputPDF, 'rb'))
		pageNum = 1
		# Loop through pages
		while (pageNum <= pdfFile.numPages):
			# Get page number and text of current page
			pdfPage = pdfFile.getPage(pageNum - 1)
			pdfText = pdfPage.extractText()
			# Get Account ID and Statement Date
			acct_id = get_acct_id(pdfText)
			statement_date = get_statement_dt(pdfText)
			# Output to list
			output.append([inputPDF, pageNum, acct_id, statement_date, pdfText])
			pageNum += 1
			
	write_to_file(output)