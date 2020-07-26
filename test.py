#print("a")
import xlrd


def getDownloadFileURL(url):

	url=url.replace("https://drive.google.com/open?id=", "")
	url=url.replace("https://drive.google.com/file/d/", "")
	url=url.replace("/view?usp=sharing", "")



	fileID=url
	return "https://drive.google.com/uc?export=download&id="+fileID


def main():
	templateFile = open("template.html", "r")
	templateHTML=templateFile.read()
	templateFile.close()

	xlFile=xlrd.open_workbook("ADV.xlsx")  
	#print(xlFile.sheet_names())
	rows=295
	sheet = xlFile.sheet_by_index(0)     
	      
	# For row 0 and column 0     
	print(sheet.cell_value(0, 0))

	for row in range(1,295):
		srno=sheet.cell_value(row, 0)
		level=sheet.cell_value(row, 1)
		#print(level)
		version=sheet.cell_value(row, 2)
		classNumber=int(sheet.cell_value(row, 3)) 
		question_number=int(sheet.cell_value(row, 4)) 
		category=sheet.cell_value(row, 5)
		question_text=sheet.cell_value(row, 6)
		question_image=sheet.cell_value(row, 7)

		optionAImage=sheet.cell_value(row, 8)
		optionAText=sheet.cell_value(row, 9)

		optionBImage=sheet.cell_value(row, 10)
		optionBText=sheet.cell_value(row, 11)

		optionCImage=sheet.cell_value(row, 12)
		optionCText=sheet.cell_value(row, 13)

		optionDImage=sheet.cell_value(row, 14)
		optionDText=sheet.cell_value(row, 15)
		answer=sheet.cell_value(row, 16)

		

		outputfileName=str(row)+".html"
		previousfileName=str(row-1)+".html"
		nextFileName=str(row+1)+".html"
		outfile=open(outputfileName, 'w')
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
		outfile.write(htmlData)
		outfile.close()
	

main()
		

	





