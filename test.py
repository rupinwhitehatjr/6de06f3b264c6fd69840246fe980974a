#print("a")
import xlrd
import os


def getDownloadFileURL(url):

	url=url.replace("https://drive.google.com/open?id=", "")
	url=url.replace("https://drive.google.com/file/d/", "")
	url=url.replace("/view?usp=sharing", "")



	fileID=url
	return "https://drive.google.com/uc?export=download&id="+fileID
def formatText(textData):
	
	textData=str(textData)	
	if(textData.isnumeric()):
		textData=int(textData)
	textData=textData.strip()
	
	
	return textData


def main():
	xlFile=xlrd.open_workbook("QuizQuestions.xlsx")
	indexTemplatefile = open("indexTemplate.html", "r")
	indexhtml=indexTemplatefile.read()
	indexTemplatefile.close()

	with open('index.html', 'wb') as temp_file:
			temp_file.write(bytes(indexhtml, 'utf-8'))
	
	#templateFile = open("template.html", "r")
	#templateHTML=templateFile.read()
	index=0
	for sheetname in xlFile.sheet_names():
		if not os.path.exists(sheetname):
			os.mkdir(sheetname)
		generateHTMLFiles(xlFile,index,sheetname)
		updateIndexFile(sheetname)
		index=index+1

def updateIndexFile(lsheetname):
	indexfile = open("index.html", "r")
	indexhtml=indexfile.read()
	indexfile.close()

	buttonTemplate = open("buttonTemplate.html", "r")
	buttonHTML=buttonTemplate.read()
	buttonTemplate.close()


	buttonHTML=buttonHTML.replace("#link", lsheetname+'/1.html')
	buttonHTML=buttonHTML.replace("#sheetname", lsheetname)

	newindexhtml=indexhtml.replace("#nextbutton",buttonHTML)
	with open('index.html', 'wb') as temp_file:
			temp_file.write(bytes(newindexhtml, 'utf-8'))




def generateHTMLFiles(workbook, sheetIndex,foldername):
	templateFile = open("template.html", "r")
	templateHTML=templateFile.read()
	templateFile.close()

	
	sheet = workbook.sheet_by_index(sheetIndex)     
	rows=sheet.nrows
	#print(rows)
	# For row 0 and column 0     
	#print(sheet.cell_value(0, 0))

	for row in range(1,rows):
		srno=formatText(sheet.cell_value(row, 0))
		level=formatText(sheet.cell_value(row, 1))
		print(row)
		
		version=formatText(sheet.cell_value(row, 2))
		classNumber=formatText(sheet.cell_value(row, 3)) 
		question_number=formatText(sheet.cell_value(row, 4)) 
		category=formatText(sheet.cell_value(row, 5))
		question_text=formatText(sheet.cell_value(row, 6))
		question_image=formatText(sheet.cell_value(row, 7))

		optionAImage=sheet.cell_value(row, 8)
		optionAText=formatText(sheet.cell_value(row, 9))

		optionBImage=sheet.cell_value(row, 10)
		optionBText=formatText(sheet.cell_value(row, 11))

		optionCImage=sheet.cell_value(row, 12)
		optionCText=formatText(sheet.cell_value(row, 13))

		optionDImage=sheet.cell_value(row, 14)
		optionDText=formatText(sheet.cell_value(row, 15))
		answer=sheet.cell_value(row, 16)
		explaination=sheet.cell_value(row, 19)

		

		outputfileName=str(row)+".html"
		if(row!=1):
			previousfileName=str(row-1)+".html"
		else:
			previousfileName="index.html"	
		nextFileName=str(row+1)+".html"
		#outfile=open(outputfileName, 'w')
		htmlData=templateHTML

		htmlData=htmlData.replace("#QuestionNumber", str(question_number))
		htmlData=htmlData.replace("#class", str(classNumber))
		htmlData=htmlData.replace("#version", str(version))
		
		htmlData=htmlData.replace("#Level", level)
		htmlData=htmlData.replace("#questionimage", "<img src='"+getDownloadFileURL(question_image)+"'>")
		
		#htmlData=htmlData.replace("#optionA", "<img src='"+optionAImage+"'>")
		if(optionAText):
			htmlData=htmlData.replace("#optionA", optionAText)
			htmlData=htmlData.replace("#optionB", optionBText)
			htmlData=htmlData.replace("#optionC", optionCText)
			htmlData=htmlData.replace("#optionD", optionDText)
		else:
			htmlData=htmlData.replace("#optionA", "<img src='"+getDownloadFileURL(optionAImage)+"'>")
			htmlData=htmlData.replace("#optionB", "<img src='"+getDownloadFileURL(optionBImage)+"'>")
			htmlData=htmlData.replace("#optionC", "<img src='"+getDownloadFileURL(optionCImage)+"'>")
			htmlData=htmlData.replace("#optionD", "<img src='"+getDownloadFileURL(optionDImage)+"'>")	
		htmlData=htmlData.replace("#question", question_text)
		htmlData=htmlData.replace("#AnswerOption", answer)
		htmlData=htmlData.replace("#previousLink", previousfileName)
		htmlData=htmlData.replace("#nextLink", nextFileName)
		htmlData=htmlData.replace("#Explaination", str(explaination))
		with open(os.path.join(foldername, outputfileName), 'wb') as temp_file:
			temp_file.write(bytes(htmlData, 'utf-8'))
		#outfile.write(htmlData)
		#outfile.close()
	

main()
		

	





