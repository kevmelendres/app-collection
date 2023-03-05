
from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox, QFileDialog
import xlwings as xw
import os.path

class FileManagement():

	@staticmethod
	def saveFile(self):
		# options = QFileDialog.Options()
		# options |= QFileDialog.DontUseNativeDialog
		fileName, _ = QFileDialog.getSaveFileName(self,"Save File As","","CTD Files (*.ctd)")
		
		if fileName:

			savedFile= open(fileName,"w+")
			savedFile.write(self.inputDesignation.text()+"\n")
			savedFile.write(self.inputFloorLevel.text()+"\n")
			savedFile.write(self.inputDetailingType.text()+"\n")
			savedFile.write(self.inputWidth.text()+"\n")
			savedFile.write(self.inputDepth.text()+"\n")
			savedFile.write(self.inputConcCover.text()+"\n")
			savedFile.write(self.inputAxialForce.text()+"\n")
			savedFile.write(self.inputFc.text()+"\n")
			savedFile.write(self.inputJTfy.currentText()+"\n")
			savedFile.write(self.inputTieFy.currentText()+"\n")
			savedFile.write(self.inputNumLongBar.text()+"\n")
			savedFile.write(self.inputLongBarDia.text()+"\n")
			savedFile.write(self.inputLongFy.currentText()+"\n")
			savedFile.write(self.label_designResults.text()+"\n")



	@staticmethod
	def loadFile(self):
						
		filePath, _ = QFileDialog.getOpenFileName(self,"Select a File",'.',"CTD (*.ctd)")

		if filePath:

			loadFile= open(filePath).readlines()
			totalLines = len(loadFile)

			self.inputDesignation.setText(loadFile[0].strip())
			self.inputFloorLevel.setText(loadFile[1].strip())
			self.inputDetailingType.setText(loadFile[2].strip())
			self.inputWidth.setText(loadFile[3].strip())
			self.inputDepth.setText(loadFile[4].strip())
			self.inputConcCover.setText(loadFile[5].strip())
			self.inputAxialForce.setText(loadFile[6].strip())
			self.inputFc.setText(loadFile[7].strip())
			self.inputJTfy.setCurrentText(loadFile[8].strip())
			self.inputTieFy.setCurrentText(loadFile[9].strip())
			self.inputNumLongBar.setText(loadFile[10].strip())
			self.inputLongBarDia.setText(loadFile[11].strip())
			self.inputLongFy.setCurrentText(loadFile[12].strip())

			resultText = ""
						
			for i in range(13,totalLines):		
				resultText += loadFile[i]

			self.label_designResults.setText(resultText)
			
	@staticmethod
	def printTemplate(self):

		if self.btnOpenTemplate.text() != "Open Template":

			isTemplateInDir = os.path.isfile('./ColumnTransverseDesignerTemplate.xlsx')
				
			if isTemplateInDir == True:
				self.templatePath = 'ColumnTransverseDesignerTemplate.xlsx'
	
			else:
				msgBox = QMessageBox()
				msgBox.setIcon(QMessageBox.Information)
				msgBox.setText("Cannot find template in default directory. Browse for template?")
				msgBox.setWindowTitle("Error")
				msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.Cancel)
				respoBrowse = msgBox.exec()

				if respoBrowse == QMessageBox.Yes:
					templatePath, _ = QFileDialog.getOpenFileName(self,"Select a File",'.',"Excel Files (*.xlsx *.xlsm)")
					self.templatePath = templatePath

					if templatePath:
							
						msgOpenTemp = QMessageBox()
						msgOpenTemp.setIcon(QMessageBox.Information)
						msgOpenTemp.setText("Template successfully selected. Opening template.")
						msgOpenTemp.setWindowTitle("Template Selected!")
						msgOpenTemp.setStandardButtons(QMessageBox.Ok)
						respoValOpen = msgOpenTemp.exec()
					
						# if respoValOpen == QMessageBox.Yes:
						# 	excel_book = xw.App().books.open(self.templatePath,update_links=False, read_only=False, ignore_read_only_recommended=True,notify=None)
						
						FileManagement.saveTemplate(self)

			if self.templatePath != "":
				self.btnOpenTemplate.setText("Open Template")

		else:
			FileManagement.saveTemplate(self)


	@staticmethod
	def saveTemplate(self):
		wb = xw.Book(self.templatePath)
		ws = wb.sheets['SPREADSHEET']

		# options = QFileDialog.Options()
		# options |= QFileDialog.DontUseNativeDialog
		filePath, _ = QFileDialog.getSaveFileName(self,"Save File As","","PDF Files (*.pdf)")

		if filePath:
			pdf_path = filePath
			pdf_path = pdf_path.replace("/", "\\")

			ws.api.PageSetup.FitToPagesWide = 1
			ws.api.PageSetup.FitToPagesTall = False

			ws.range("A1:T109").api.ExportAsFixedFormat(0,pdf_path)



class BatchAnalysisFileMNGT():

	@staticmethod
	def createNewSaveFile(self):

		newFile, _ = QFileDialog.getSaveFileName(self,"Save file as",'.',"Excel File (*.xlsx)")

		if newFile:

			self.batchFilePath = newFile

			app = xw.App(visible=False)
			wb = xw.Book()
			wb.save(newFile)
			wb.close()
			
			
			msgBox = QMessageBox()
			msgBox.setIcon(QMessageBox.Information)
			msgBox.setText("All data would automatically be saved in the newly created file unless a new file is browsed. Input data in excel file for batch analysis.")
			msgBox.setWindowTitle("Information")
			msgBox.setStandardButtons(QMessageBox.Ok)
			respoBrowse = msgBox.exec()

			BatchAnalysisFileMNGT.setupNewSaveFile(self)
			excel_book = xw.App().books.open(self.batchFilePath,update_links=False, read_only=False, ignore_read_only_recommended=True,notify=None)

	def browseSaveFile(self):
		filePath, _ = QFileDialog.getOpenFileName(self,"Select a File",'.',"Excel File (*.xlsx)")
		self.batchFilePath = filePath
				
		if filePath:

			msgBox = QMessageBox()
			msgBox.setIcon(QMessageBox.Information)
			msgBox.setText("Save file selected successfully.")
			msgBox.setWindowTitle("Information")
			msgBox.setStandardButtons(QMessageBox.Ok)
			respoBrowse = msgBox.exec()

			excel_book = xw.App().books.open(self.batchFilePath,update_links=False, read_only=False, ignore_read_only_recommended=True,notify=None)



	@staticmethod
	def setupNewSaveFile(self):
		app = xw.App(visible=False)
		wb = xw.Book(self.batchFilePath)
		wb.sheets[0].name = "CONSOLIDATED"
		ws = wb.sheets["CONSOLIDATED"]
		headers = ["Column\nDesignation",
			"Floor\nLevel",
			"Type of\nDetailing",
			"Width", 
			"Depth",
			"Concrete\nCover",
			"Axial\nForce",
			"fc'",
			"Joint/\nTransverse\nfy",
			"Ties\nfy",
			"Number of\nLong.Bars",
			"Long. Bar\nDiameter",
			"Steel Strength\nfy"]

		ws.range("N1").value = "Input Errors"

		for i in range(len(headers)):
			ws.range(1,i+1).value = headers[i]

		headerResults = ["njo",
			"N-Outer\nx",
			"N-Outer\ny",
			"N-Inner\nx", 
			"N-Inner\ny",
			"djo",
			"dji",
			"sj",
			"dco",
			"dci",
			"sc",
			"dto",
			"dti",
			"st",
			]

		for i in range(len(headerResults)):
			ws.range(1,i+16).value = headerResults[i]

		
		ws.range("A:AF").column_width = 13.5

		wb.save(self.batchFilePath)
		app.quit()





		

