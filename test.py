#print("a")
import xlrd
import os
from datetime import datetime
allowed=['BEG (Evaluation)', 'INT(Evaluation)', 'Advanced Quiz Questions', 'PRO(Evaluation)', 'PRO(V2)']
def getDownloadFileURL(url):
	isdrawing=False
	if("drawing" in url):
		#print(url)
		isdrawing=True

	url=url.replace("https://drive.google.com/open?id=", "")
	url=url.replace("https://docs.google.com/drawings/d/", "")
	url=url.replace("https://drive.google.com/file/d/", "")
	url=url.replace("/edit?usp=sharing", "")	
	url=url.replace("/view?usp=sharing", "")
	if(url.strip()+""==""):
		downloadURL=""
	else:
		if(isdrawing):
			downloadURL="https://docs.google.com/drawings/d/"+url+"/export/png"
		else:	
			downloadURL="https://drive.google.com/uc?export=download&id="+url
	
	return downloadURL

def formatText(textData):
	
	textData=str(textData)
	textData=textData.strip()
	if(textData.isnumeric()):
		textData=int(textData)
	
	
	
	return str(textData)


def main():
	xlFile=xlrd.open_workbook("Be a Quiz Master - Challenge (Responses).xlsx")
	indexTemplatefile = open("indexTemplate.html", "r")
	indexhtml=indexTemplatefile.read()
	indexTemplatefile.close()

	with open('index.html', 'wb') as temp_file:
			temp_file.write(bytes(indexhtml, 'utf-8'))
	
	#templateFile = open("template.html", "r")
	#templateHTML=templateFile.read()
	index=0
	#print(xlFile.sheet_names())
	#exit()
	for sheetname in xlFile.sheet_names():
		
		if(sheetname in allowed):
			print(sheetname)
			if not os.path.exists(sheetname):
				os.mkdir(sheetname)
			generateHTMLFiles(xlFile,index,sheetname)
			updateIndexFile(sheetname)
		index=index+1
	addDateInIndexFile()

def addDateInIndexFile():
	indexfile = open("index.html", "r")
	indexhtml=indexfile.read()
	indexfile.close()

	

	today = datetime.now()

	# dd/mm/YY
	d1 = today.strftime("%d/%m/%Y %T")	
	#print(d1)

	indexhtml=indexhtml.replace("#nextbutton","&nbsp;")
	indexhtml=indexhtml.replace("#date",d1)
	with open('index.html', 'wb') as temp_file:
			temp_file.write(bytes(indexhtml, 'utf-8'))


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
	offset=2
	for row in range(1,rows):
		srno=formatText(sheet.cell_value(row, offset+0))
		level=formatText(sheet.cell_value(row, offset+1))
		#print(row)
		
		version=formatText(sheet.cell_value(row, offset+2))
		classNumber=formatText(sheet.cell_value(row, offset+3)) 
		question_number=formatText(sheet.cell_value(row, offset+4)) 
		category=formatText(sheet.cell_value(row, offset+5))
		question_text=formatText(sheet.cell_value(row, offset+6))
		question_image=formatText(sheet.cell_value(row, offset+7))
		#print(question_text)

		optionAImage=sheet.cell_value(row, offset+8)
		optionAText=formatText(sheet.cell_value(row, offset+9))

		optionBImage=sheet.cell_value(row, offset+10)
		optionBText=formatText(sheet.cell_value(row, offset+11))

		optionCImage=sheet.cell_value(row, offset+12)
		optionCText=formatText(sheet.cell_value(row, offset+13))

		optionDImage=sheet.cell_value(row, offset+14)
		optionDText=formatText(sheet.cell_value(row, offset+15))
		answer=sheet.cell_value(row, offset+16)
		explaination=sheet.cell_value(row, offset+17)
		activity=sheet.cell_value(row, offset+18)
		solution=sheet.cell_value(row, offset+19)

		

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
		imagePath=getDownloadFileURL(question_image)
		if(imagePath):
			htmlData=htmlData.replace("#questionimage", "<img src='"+imagePath+"'>")
		
		#htmlData=htmlData.replace("#optionA", "<img src='"+optionAImage+"'>")
		if(optionAText):
			htmlData=htmlData.replace("#optionA", optionAText)
			
		else:
			htmlData=htmlData.replace("#optionA", "<img src='"+getDownloadFileURL(optionAImage)+"'>")
			

		if(optionBText):
			
			htmlData=htmlData.replace("#optionB", optionBText)
			
			
		else:
			
			htmlData=htmlData.replace("#optionB", "<img src='"+getDownloadFileURL(optionBImage)+"'>")
			
		if(optionCText):
			
			htmlData=htmlData.replace("#optionC", optionCText)
			
		else:
			
			htmlData=htmlData.replace("#optionC", "<img src='"+getDownloadFileURL(optionCImage)+"'>")
				
		if(optionDText):
			
			htmlData=htmlData.replace("#optionD", optionDText)
		else:
			
			htmlData=htmlData.replace("#optionD", "<img src='"+getDownloadFileURL(optionDImage)+"'>")	

		htmlData=htmlData.replace("#question", question_text)
		htmlData=htmlData.replace("#AnswerOption", answer)
		htmlData=htmlData.replace("#previousLink", previousfileName)
		htmlData=htmlData.replace("#nextLink", nextFileName)
		htmlData=htmlData.replace("#Explaination", str(explaination))
		htmlData=htmlData.replace("#activity", str(activity))
		htmlData=htmlData.replace("#solution", str(solution))
		with open(os.path.join(foldername, outputfileName), 'wb') as temp_file:
			temp_file.write(bytes(htmlData, 'utf-8'))
		#outfile.write(htmlData)
		#outfile.close()
	

main()
		

	





