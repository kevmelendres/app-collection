from ApplicationCollectionCalculationModule import ColumnTransverseDesignerCalculations as ctdCalc, BatchAnalysisMethods as BAM
from FileManagement import FileManagement as fmgt, BatchAnalysisFileMNGT as BatchMNGT


from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox, QFileDialog
from PyQt5.uic import loadUi
import sys
import xlwings as xw
import os.path


class MainApplication(QMainWindow):
	def __init__(self):
		super().__init__()
		self.ui = loadUi("ApplicationCollection.ui",self)
		self.setFixedSize(688, 587)		
		self.btnColTransverseDesigner.clicked.connect(self.openColumnTransVerseDesignerWindow)
		



	def openColumnTransVerseDesignerWindow(self):
		self.hide()
		openWindow = ColumnTransverseDesigner()

	def closeEvent(self, event):
	 	msgBox = QMessageBox()
	 	msgBox.setIcon(QMessageBox.Warning)
	 	msgBox.setText("Are you sure you want to quit?")
	 	msgBox.setWindowTitle("Quit?")
	 	msgBox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
	 	responseVal = msgBox.exec()
	 	if responseVal == QMessageBox.Ok:
	 		event.accept()
	 	else:
	 		event.ignore()

		
class ColumnTransverseDesigner(QMainWindow):

	def __init__(self):
		super().__init__()
		self.ui = loadUi("ColumnTransverseDesigner.ui",self)
		self.templatePath = ""
		self.batchFilePath = ""
		self.setFixedSize(791, 674)	
		self.show()
		self.btnOpenTemplate.clicked.connect(self.openCTDTemplate)
		self.btnDesign.clicked.connect(self.designTransverseReinf)
		self.mtemplate_browse.triggered.connect(self.browseCTDTemplate)
		self.mtemplate_open.triggered.connect(self.openCTDTemplate)
		self.mfile_load.triggered.connect(lambda: fmgt.loadFile(self))
		self.mfile_save.triggered.connect(lambda: fmgt.saveFile(self))
		self.mtemplate_print.triggered.connect(lambda: fmgt.printTemplate(self))
		self.mbatch_new.triggered.connect(lambda: BatchMNGT.createNewSaveFile(self))
		self.mbatch_browse.triggered.connect(lambda: BatchMNGT.browseSaveFile(self))
		self.mbatch_startanalysis.triggered.connect(lambda: BAM.startBatchAnalysis(self))


		#default values
		self.inputJTfy.setCurrentText("414")
		self.inputTieFy.setCurrentText("414")
		self.inputLongFy.setCurrentText("414")


	def openCTDTemplate(self):

		if self.btnOpenTemplate.text() != "Open Template":

			isTemplateInDir = os.path.isfile('./ColumnTransverseDesignerTemplate.xlsx')
				
			if isTemplateInDir == True:
				msgTempInDir = QMessageBox()
				msgTempInDir.setIcon(QMessageBox.Information)
				msgTempInDir.setText("Template found in the same directory. Would you like to open it?")
				msgTempInDir.setWindowTitle("Template Found!")
				msgTempInDir.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
				respoVal = msgTempInDir.exec()

				self.templatePath = 'ColumnTransverseDesignerTemplate.xlsx'

				if respoVal == QMessageBox.Yes:
					excel_book = xw.App().books.open(self.templatePath,update_links=False, read_only=False, ignore_read_only_recommended=True,notify=None)
				

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

					print(templatePath)

					if templatePath != "":
						msgOpenTemp = QMessageBox()
						msgOpenTemp.setIcon(QMessageBox.Information)
						msgOpenTemp.setText("Template successfully selected. Open template?")
						msgOpenTemp.setWindowTitle("Template Selected!")
						msgOpenTemp.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
						respoValOpen = msgOpenTemp.exec()
					
					if respoValOpen == QMessageBox.Yes:
						excel_book = xw.App().books.open(self.templatePath,update_links=False, read_only=False, ignore_read_only_recommended=True,notify=None)
				
			if self.templatePath != "":
				self.btnOpenTemplate.setText("Open Template")
		else:
			excel_book = xw.App().books.open(self.templatePath,update_links=False, read_only=False, ignore_read_only_recommended=True,notify=None)



	def designTransverseReinf(self):

		self.label_designResults.setText("")
		
		QApplication.processEvents()

		if self.btnOpenTemplate.text() == "Find Template":

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
										
					if templatePath != "":
						
						msgOpenTemp = QMessageBox()
						msgOpenTemp.setIcon(QMessageBox.Information)
						msgOpenTemp.setText("Template successfully selected. Opening template.")
						msgOpenTemp.setWindowTitle("Template Selected!")
						msgOpenTemp.setStandardButtons(QMessageBox.Ok)
						respoValOpen = msgOpenTemp.exec()

						self.btnOpenTemplate.setText("Open Template")
				
						# if respoValOpen == QMessageBox.Yes:
						# 	excel_book = xw.App().books.open(self.templatePath,update_links=False, read_only=False, ignore_read_only_recommended=True,notify=None)

						self.statusLabel.setText("Calculating design. Please wait...")
						QApplication.processEvents()

						ctdCalc.CTDinputInitialData(self)

						if self.totalError == 0:
							ctdCalc.CTDJointConfinementDesign(self)
							ctdCalc.CTDTieDesign(self)
							ctdCalc.writeResultsToUI(self)

						self.statusLabel.setText("Calculations complete!")
						QApplication.processEvents()
			
		else:
			self.statusLabel.setText("Calculating design. Please wait...")
			QApplication.processEvents()
			
			ctdCalc.CTDinputInitialData(self)

			if self.totalError == 0:
				ctdCalc.CTDJointConfinementDesign(self)
				ctdCalc.CTDTieDesign(self)
				ctdCalc.writeResultsToUI(self)

			self.statusLabel.setText("Calculations complete!")
			


	def showDialog(self):
		
		msgBox = QMessageBox()
		msgBox.setIcon(QMessageBox.Information)
		msgBox.setText("Message box pop up window")
		msgBox.setWindowTitle("QMessageBox Example")
		msgBox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
		
		returnValue = msgBox.exec()


	def browseCTDTemplate(self):

		if self.btnOpenTemplate.text() != "Open Template":

			isTemplateInDir = os.path.isfile('./ColumnTransverseDesignerTemplate.xlsx')
				
			if isTemplateInDir == True:
				msgTempInDir = QMessageBox()
				msgTempInDir.setIcon(QMessageBox.Information)
				msgTempInDir.setText("Default template found in the same directory.")
				msgTempInDir.setWindowTitle("Template Found!")
				msgTempInDir.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
				respoVal = msgTempInDir.exec()

				self.templatePath = 'ColumnTransverseDesignerTemplate.xlsx'

				if respoVal == QMessageBox.Yes:
					excel_book = xw.App().books.open(self.templatePath,update_links=False, read_only=False, ignore_read_only_recommended=True,notify=None)
				

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
						msgOpenTemp.setText("Template successfully selected. Open template?")
						msgOpenTemp.setWindowTitle("Template Selected!")
						msgOpenTemp.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
						respoValOpen = msgOpenTemp.exec()
				
						if respoValOpen == QMessageBox.Yes:
							excel_book = xw.App().books.open(self.templatePath,update_links=False, read_only=False, ignore_read_only_recommended=True,notify=None)
				
			if self.templatePath != "":
				self.btnOpenTemplate.setText("Open Template")
		else:

			msgTempOverride = QMessageBox()
			msgTempOverride.setIcon(QMessageBox.Warning)
			msgTempOverride.setText("Template already selected. Select another template?")
			msgTempOverride.setWindowTitle("Template Already Selected!")
			msgTempOverride.setStandardButtons(QMessageBox.Yes | QMessageBox.No)


			respoVal = msgTempOverride.exec()

			if respoVal == QMessageBox.Yes:
				templatePath, _ = QFileDialog.getOpenFileName(self,"Select a File",'.',"Excel Files (*.xlsx *.xlsm)")
				self.templatePath = templatePath
				self.btnOpenTemplate.setText("Open Template")

	
   		
	def closeEvent(self, event):
	 	mainWindow = MainApplication()
	 	mainWindow.show()






if __name__=="__main__":
	app = QApplication(sys.argv)
	ui = MainApplication()
	ui.show()
	sys.exit(app.exec_())



