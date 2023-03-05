import xlwings as xw
import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox, QFileDialog, QInputDialog
import os.path
# import unicodedata as ud


class ColumnTransverseDesignerCalculations():
	
	def CTDinputInitialData(self):

		inputDesignation = self.inputDesignation.text()
		inputFloorLevel = self.inputFloorLevel.text()
		inputWidth = self.inputWidth.text()
		inputDetailingType = self.inputDetailingType.text()
		inputDepth = self.inputDepth.text()
		inputConcCover = self.inputConcCover.text()
		inputAxialForce = self.inputAxialForce.text()
		inputFc = self.inputFc.text()
		inputJTfy = self.inputJTfy.currentText()
		inputTieFy = self.inputTieFy.currentText()
		inputLongFy = self.inputLongFy.currentText()
		inputNumLongBar = self.inputNumLongBar.text()
		inputLongBarDia = self.inputLongBarDia.text()

		# DATA VALIDATION

		from ApplicationCollectionCalculationModule import ColumnTransverseDesignerCalculations as ctdC

		totalError = 0
		
		validationErrorText = "WARNING:\nCheck for errors in the following data inputs:\n\n"

		if ctdC.isStringInputValid(inputDesignation) == 1:
			totalError+=1
			validationErrorText+= str(totalError) + ") " + "Designation\n"

		if ctdC.isStringInputValid(inputFloorLevel) == 1:
			totalError+=1
			validationErrorText+= str(totalError) + ") " + "Floor Level\n"

		if ctdC.isStringInputValid(inputDetailingType) == 1:
			totalError+=1
			validationErrorText+= str(totalError) + ") " + "Type of Detailing\n"

		if ctdC.isNumInputValid(inputWidth) == 1:
			totalError+=1
			validationErrorText+= str(totalError) + ") " + "Width\n"

		if ctdC.isNumInputValid(inputDepth) == 1:
			totalError+=1
			validationErrorText+= str(totalError) + ") " + "Depth\n"

		if ctdC.isNumInputValid(inputConcCover) == 1:
			totalError+=1
			validationErrorText+= str(totalError) + ") " + "Concrete Cover\n"

		if ctdC.isNumInputValid(inputAxialForce) == 1:
			totalError+=1
			validationErrorText+= str(totalError) + ") " + "Axial Force\n"

		if ctdC.isNumInputValid(inputFc) == 1:
			totalError+=1
			validationErrorText+= str(totalError) + ") " + "Concrete fc'\n"

		if ctdC.isNumInputValid(inputNumLongBar) == 1:
			totalError+=1
			validationErrorText+= str(totalError) + ") " + "Number of Long. Bars\n"

		if ctdC.isNumInputValid(inputLongBarDia) == 1:
			totalError+=1
			validationErrorText+= str(totalError) + ") " + "Long. Bar Diameter\n"

		validationErrorText = validationErrorText.rstrip('\n')
		validationErrorText+= "\n\nProcess will not continue."
		self.totalError = totalError

		# DATA VALIDATION

		

		if totalError > 0:
			msgBox = QMessageBox()
			msgBox.setIcon(QMessageBox.Warning)
			msgBox.setText(validationErrorText)
			msgBox.setWindowTitle("Check Input Errors")
			msgBox.setStandardButtons(QMessageBox.Ok)
			returnVal = msgBox.exec()

		else:

			# wb = xw.Book('ColumnTransverseDesignerTemplate.xlsx')
			wb = xw.Book(self.templatePath)
			ws = wb.sheets['SPREADSHEET']
			ws.range("F10").value = inputDesignation
			ws.range("F11").value = inputFloorLevel
			ws.range("F12").value = inputDetailingType
			ws.range("F16").value = "Rectangular"
			ws.range("F17").value = inputWidth
			ws.range("F18").value = inputDepth
			ws.range("F19").value = inputConcCover
			ws.range("F23").value = inputAxialForce
			ws.range("O30").value = inputFc
			ws.range("O33").value = inputLongFy
			ws.range("O34").value = inputJTfy
			ws.range("O35").value = inputJTfy
			ws.range("Q34").value = inputJTfy
			ws.range("Q35").value = inputJTfy
			ws.range("O36").value = inputTieFy
			ws.range("Q36").value = inputTieFy
			ws.range("O41").value = inputNumLongBar
			ws.range("O42").value = inputLongBarDia
			ws.range("O45").value = 2
			ws.range("Q45").value = 2
			ws.range("O46").value = 0
			ws.range("Q46").value = 0



	@staticmethod
	def isStringInputValid(data):
		returnVal = 0
		
		if data == "":
			returnVal = 1

		return returnVal


	@staticmethod
	def isNumInputValid(data):
		returnVal = 0

		try:
			float(data)
		except:
			TypeError
			returnVal = 1

		if data == "":
			returnVal = 1

		return returnVal

	






	@staticmethod	
	def CTDTieDesign(self):

		allowedDCR = 1.00

		spacingIncrement = 25
		useTieDia = [12,16,20,25]
		diaCtr = 0
		spacingCtr = 300
		endSpacing = 100

		# wb = xw.Book('ColumnTransverseDesignerTemplate.xlsx')
		wb = xw.Book(self.templatePath)

		ws = wb.sheets['SPREADSHEET']

		while diaCtr < len(useTieDia):

			spacingCtr = 300

			while spacingCtr>=endSpacing:

				ws.range("O61").value = useTieDia[diaCtr]
				ws.range("O62").value = useTieDia[diaCtr]
				ws.range("O63").value = spacingCtr

				OKValue = str(ws.range("O83").value)
				xDCR = ws.range("O70").value
				yDCR = ws.range("Q70").value
				isTieSpacingOK = str(ws.range("O83").value)


				if OKValue == "OK" and xDCR<allowedDCR and yDCR<allowedDCR and isTieSpacingOK == "OK":
					break

				spacingCtr-=spacingIncrement


			if OKValue == "OK" and xDCR<allowedDCR and yDCR<allowedDCR and isTieSpacingOK == "OK":
				break

			diaCtr+=1

	@staticmethod	
	def CTDJointConfinementDesign(self):

		spacingIncrement = 25
		useTieDia = [12,16,25,28]

		diaCtr = 0
		spacingCtr = 300
		endSpacing = 100

		wb = xw.Book(self.templatePath)
		ws = wb.sheets['SPREADSHEET']

		while diaCtr < len(useTieDia):

		
			spacingCtr = 300

			ws.range("O47").value = useTieDia[diaCtr]
			ws.range("O48").value = useTieDia[diaCtr]
			ws.range("O46").value = 0
			ws.range("Q46").value = 0

			while spacingCtr>=endSpacing:				
				ws.range("O49").value = spacingCtr
				OKValueJoint = str(ws.range('O81').value)
				OKValueConfinement = str(ws.range('O82').value)

				reqdJointSpacing = ws.range('O94').value
				actJointSpacing = ws.range('O49').value

				reqdConfinSpacing = ws.range('O103').value
				actConfinSpacing = ws.range('O56').value
				
				if OKValueJoint == "OK" and OKValueConfinement == "OK":
				
					break

				spacingCtr-=spacingIncrement



			
			ColumnTransverseDesignerCalculations.CTDMinNumLegsX(self)
			ColumnTransverseDesignerCalculations.CTDMinNumLegsY(self)

			isMinLongBarSpacingOK = str(ws.range('O78').value)

			OKValueJoint = str(ws.range('O81').value)
			OKValueConfinement = str(ws.range('O82').value)
		
	
			if OKValueJoint == "OK" and OKValueConfinement == "OK" and isMinLongBarSpacingOK == "OK":

				break
			else:
				diaCtr+=1
				
			
	@staticmethod	
	def CTDMinNumLegsX(self):

		legXCtr = 0
		allowedDCR = 1.00

		# wb = xw.Book('ColumnTransverseDesignerTemplate.xlsx')
		wb = xw.Book(self.templatePath)
		ws = wb.sheets['SPREADSHEET']

		DCXJoint = ws.range("O68").value 
		DCXConfinement = ws.range("O69").value 
		ws.range("O46").value = legXCtr
		
		while DCXJoint > allowedDCR or DCXConfinement > allowedDCR:
			ws.range("O46").value = legXCtr
			DCXJoint = ws.range("O68").value 
			DCXConfinement = ws.range("O69").value
			legXCtr+=1


		maxXSpacing = ws.range("W79").value 
		actualXSpacing = ws.range("U79").value 
		
		while actualXSpacing > maxXSpacing:
			ws.range("O46").value = legXCtr
			maxXSpacing = ws.range("W79").value 
			actualXSpacing = ws.range("U79").value 
			legXCtr+=1

	@staticmethod	
	def CTDMinNumLegsY(self):

		legYCtr = 0
		allowedDCR = 1.00

		# wb = xw.Book('ColumnTransverseDesignerTemplate.xlsx')
		wb = xw.Book(self.templatePath)
		ws = wb.sheets['SPREADSHEET']

		DCYJoint = ws.range("Q68").value 
		DCYConfinement = ws.range("Q69").value 
		ws.range("Q46").value = legYCtr
		
		while DCYJoint > allowedDCR or DCYConfinement > allowedDCR:
			ws.range("Q46").value = legYCtr
			DCYJoint = ws.range("Q68").value 
			DCYConfinement = ws.range("Q69").value
			legYCtr+=1

		maxYSpacing = ws.range("W79").value 
		actualYSpacing = ws.range("V79").value 
		
		while actualYSpacing > maxYSpacing:
			ws.range("Q46").value = legYCtr
			maxYSpacing = ws.range("W79").value 
			actualYSpacing = ws.range("V79").value 
			legYCtr+=1

	@staticmethod	
	def writeResultsToUI(self):

		wb = xw.Book(self.templatePath)
		ws = wb.sheets["SPREADSHEET"]

		# diaSym = ud.lookup(eta)

		rep0 = "DESIGN REQUIREMENTS:\n"
		rep1 = "No. of outer legs along X:\t" + str(int(ws.range("O45").value)) + "\n"
		rep2 = "No. of outer legs along Y:\t" + str(int(ws.range("Q45").value)) + "\n"
		rep3 = "No. of inner legs along X:\t" + str(int(ws.range("O46").value)) + "\n"
		rep4 = "No. of inner legs along Y:\t" + str(int(ws.range("Q46").value)) + "\n"
		rep5 = "\nDUCTILITY REQUIREMENTS:\n"
		rep6 = "Joint Reinforcement:\tUse ø" + str(int(ws.range("O47").value)) + " @ "+str(int(ws.range("O49").value)) + "\n"
		rep7 = "Conf. Zone Reinforcement:\tUse ø" + str(int(ws.range("O54").value)) + " @ "+str(int(ws.range("O56").value)) + "\n"
		rep8 = "Tie Zone Reinforcement:\tUse ø" + str(int(ws.range("O61").value)) + " @ "+str(int(ws.range("O63").value))

		otherRmkCtr = 1

		if str(ws.range("O75").value) == "NOT OK" or str(ws.range("O76").value) == "NOT OK" or str(ws.range("O77").value) == "NOT OK":
			rep9 = "\n\nOTHER REMARKS:\n"
		else:
			rep9 = ""

		if str(ws.range("O75").value) == "NOT OK":
			rep9 += str(otherRmkCtr) + ". Increase number of bars to satisfy ductility requirements.\n"
			otherRmkCtr+=1

		if str(ws.range("O76").value) == "NOT OK":
			rep9 += str(otherRmkCtr) + ". Steel ratio < 1%. Increase bars.\n"
			otherRmkCtr+=1

		if str(ws.range("O77").value) == "NOT OK":
			rep9 += str(otherRmkCtr) + ". Steel ratio > 6%. Revise design.\n"
			otherRmkCtr+=1

		rep9 = rep9[:-1]



		finalReport = rep0 + rep1 + rep2 + rep3+ rep4 + rep5 + rep6 + rep7 + rep8 + rep9

		self.label_designResults.setText(finalReport)



class BatchAnalysisMethods():

	def ValidateDataRow(self, rowNum):

		wb = xw.Book(self.batchFilePath)
		ws = wb.sheets["CONSOLIDATED"]

		inputDesignation = ws.range("A"+str(rowNum)).value
		inputFloorLevel = ws.range("B"+str(rowNum)).value
		inputDetailingType = ws.range("C"+str(rowNum)).value
		inputWidth = ws.range("D"+str(rowNum)).value
		inputDepth = ws.range("E"+str(rowNum)).value
		inputConcCover = ws.range("F"+str(rowNum)).value
		inputAxialForce = ws.range("G"+str(rowNum)).value
		inputFc = ws.range("H"+str(rowNum)).value
		inputJTfy = ws.range("I"+str(rowNum)).value
		inputTieFy = ws.range("J"+str(rowNum)).value
		inputNumLongBar = ws.range("K"+str(rowNum)).value
		inputLongBarDia = ws.range("L"+str(rowNum)).value
		inputLongFy = ws.range("M"+str(rowNum)).value

		# DATA VALIDATION

		from ApplicationCollectionCalculationModule import ColumnTransverseDesignerCalculations as ctdC

		totalError = 0
		
		validationErrorText = ""

		if ctdC.isStringInputValid(inputDesignation) == 1:
			totalError+=1
			validationErrorText+= "Designation, "

		if ctdC.isStringInputValid(inputFloorLevel) == 1:
			totalError+=1
			validationErrorText+= "Floor Level, "

		if ctdC.isStringInputValid(inputDetailingType) == 1:
			totalError+=1
			validationErrorText+= "Type of Detailing, "

		if ctdC.isNumInputValid(inputWidth) == 1:
			totalError+=1
			validationErrorText+= "Width, "

		if ctdC.isNumInputValid(inputDepth) == 1:
			totalError+=1
			validationErrorText+= "Depth, "

		if ctdC.isNumInputValid(inputConcCover) == 1:
			totalError+=1
			validationErrorText+= "Concrete Cover, "

		if ctdC.isNumInputValid(inputAxialForce) == 1:
			totalError+=1
			validationErrorText+= "Axial Force, "

		if ctdC.isNumInputValid(inputFc) == 1:
			totalError+=1
			validationErrorText+= "Concrete fc', "

		if ctdC.isNumInputValid(inputNumLongBar) == 1:
			totalError+=1
			validationErrorText+= "Number of Long. Bars, "

		if ctdC.isNumInputValid(inputLongBarDia) == 1:
			totalError+=1
			validationErrorText+= "Long. Bar Diameter, "

		if totalError>0:
			validationErrorText = validationErrorText[:-2]

		
		return totalError,validationErrorText
				
	
	def InputDataToForm(self, rowNum):

		wb = xw.Book(self.batchFilePath)
		ws = wb.sheets["CONSOLIDATED"]

		self.inputDesignation.setText(str(ws.range("A"+str(rowNum)).value))
		self.inputFloorLevel.setText(str(ws.range("B"+str(rowNum)).value))
		self.inputWidth.setText(str(ws.range("D"+str(rowNum)).value))
		self.inputDetailingType.setText(str(ws.range("C"+str(rowNum)).value))
		self.inputDepth.setText(str(ws.range("E"+str(rowNum)).value))
		self.inputConcCover.setText(str(ws.range("F"+str(rowNum)).value))
		self.inputAxialForce.setText(str(ws.range("G"+str(rowNum)).value))
		self.inputFc.setText(str(ws.range("H"+str(rowNum)).value))
		self.inputJTfy.setCurrentText(str(ws.range("I"+str(rowNum)).value))
		self.inputTieFy.setCurrentText(str(ws.range("J"+str(rowNum)).value))
		self.inputLongFy.setCurrentText(str(ws.range("M"+str(rowNum)).value))
		self.inputNumLongBar.setText(str(ws.range("K"+str(rowNum)).value))
		self.inputLongBarDia.setText(str(ws.range("L"+str(rowNum)).value))

	
	def WriteResultsToWS(self,rowNum):

		wbTo = xw.Book(self.batchFilePath)
		wsTo = wbTo.sheets["CONSOLIDATED"]

		wbFrom = xw.Book(self.templatePath)
		wsFrom = wbFrom.sheets["SPREADSHEET"]

		wsTo.range("P"+str(rowNum)).value = wsFrom.range("O45").value
		wsTo.range("Q"+str(rowNum)).value = wsFrom.range("Q45").value
		wsTo.range("R"+str(rowNum)).value = wsFrom.range("O46").value
		wsTo.range("S"+str(rowNum)).value = wsFrom.range("Q46").value
		wsTo.range("T"+str(rowNum)).value = wsFrom.range("O47").value
		wsTo.range("U"+str(rowNum)).value = wsFrom.range("O48").value
		wsTo.range("V"+str(rowNum)).value = wsFrom.range("O49").value
		wsTo.range("W"+str(rowNum)).value = wsFrom.range("O54").value
		wsTo.range("X"+str(rowNum)).value = wsFrom.range("O55").value
		wsTo.range("Y"+str(rowNum)).value = wsFrom.range("O56").value
		wsTo.range("Z"+str(rowNum)).value = wsFrom.range("O61").value
		wsTo.range("AA"+str(rowNum)).value = wsFrom.range("O62").value
		wsTo.range("AB"+str(rowNum)).value = wsFrom.range("O63").value







	@staticmethod
	def startBatchAnalysis(self):

		from BaseFile_ApplicationCollection import ColumnTransverseDesigner as BFAC
		from ApplicationCollectionCalculationModule import ColumnTransverseDesignerCalculations as ctdCalc
		from FileManagement import BatchAnalysisFileMNGT as BAFM

		print(self.batchFilePath)
		print(self.templatePath)

		if self.templatePath == "":
			BFAC.openCTDTemplate(self)

		if self.batchFilePath == "":
			msgBox = QMessageBox()
			msgBox.setIcon(QMessageBox.Information)
			msgBox.setText("No save file selected. Browse for selected file.")
			msgBox.setWindowTitle("Error")
			msgBox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
			respoBrowse = msgBox.exec()

			if respoBrowse == QMessageBox.Ok:
				BAFM.browseSaveFile(self)


		if self.templatePath != "" and self.batchFilePath != "":

 			wb = xw.Book(self.batchFilePath)
 			ws = wb.sheets["CONSOLIDATED"]

 			if self.btnOpenTemplate.text() == "Find Template":

 				isTemplateInDir = os.path.isfile('./ColumnTransverseDesignerTemplate.xlsx')

 				if isTemplateInDir == True:
 					self.templatePath = 'ColumnTransverseDesignerTemplate.xlsx'

 				else:
 					msgBox = QMessageBox()
 					msgBox.setIcon(QMessageBox.Information)
 					msgBox.setText("Cannot find template in default directory. Need to import template first.")
 					msgBox.setWindowTitle("Error")
 					msgBox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
 					respoBrowse = msgBox.exec()

 					if respoBrowse == QMessageBox.Ok:
 						templatePath, _ = QFileDialog.getOpenFileName(self,"Select a File",'.',"Excel Files (*.xlsx *.xlsm)")
 						self.templatePath = templatePath

 						if templatePath != "":
 							self.btnOpenTemplate.setText("Open Template")
 							msgOpenTemp = QMessageBox()
 							msgOpenTemp.setIcon(QMessageBox.Information)
 							msgOpenTemp.setText("Template successfully selected")
 							msgOpenTemp.setWindowTitle("Template Selected!")
 							msgOpenTemp.setStandardButtons(QMessageBox.Ok)
 							respoValOpen = msgOpenTemp.exec()

 			startRow, _ = QInputDialog.getInt(self, "Input Start Row", "Input first row of data to be analyzed.")
 			endRow, _ = QInputDialog.getInt(self, "Input End Row", "Input last row of data to be analyzed.")

 			for i in range(startRow,endRow+1):
 				if BatchAnalysisMethods.ValidateDataRow(self,i)[0] != 0:
 					ws.range("N" + str(i)).value = BatchAnalysisMethods.ValidateDataRow(self,i)[1]
 					continue
 				else:
 					BatchAnalysisMethods.InputDataToForm(self,i)
 					ctdCalc.CTDJointConfinementDesign(self)
 					ctdCalc.CTDTieDesign(self)
 					BatchAnalysisMethods.WriteResultsToWS(self,i)

				

					
				



