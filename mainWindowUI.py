from PySide6.QtWidgets import (QMainWindow,QWidget,QPushButton,QVBoxLayout,QStatusBar,QLabel,QLineEdit
                               ,QGridLayout,QHBoxLayout,QSizePolicy,QDoubleSpinBox,QTimeEdit,
                               QFileDialog,QCheckBox,QScrollArea,QMessageBox,QComboBox,QTextEdit)
from PySide6.QtGui import  QIcon
from PySide6.QtCore import QSize,Qt,QTimer,QTime,Signal,QDateTime
from CheckUnfilledTeams import KPIreportVerifier,supportedExcelExtensions,runExcelMacro
import json
from datetime import datetime
import os,sys
import ast

from LogData import Database

from applicationData import (reportsAndTeamsDict,reports,reportMonth,settingsSaveFile,settingsfileName,settingsFilePath,
                       reportWeek,settingsIcon,appName,appIcon,mainWindowHeight,
                       mainWindowWidth,logFilePath,dataBaseColumnsAndDataTypes)
import calendar
from copyPasteLinksOfPDF import copyPasteLinksofPDF
import types

        

# Settings window
class settingsWidget(QWidget):
    settingsSavedSignal:Signal=Signal()
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Settings")
        self.setWindowIcon(QIcon(settingsIcon))
        self.resize(600, 600)
        
        self.reportPathHeading=QLabel("Report Excel Paths")
        self.settingsLayout=QGridLayout()
        self.setWindowIcon(QIcon(settingsIcon))

        #creating save file keys from report names
        self.pathSaveFileKeys=[]
        for report in reports:
            self.pathSaveFileKeys.append(f"{report}_Excel_Path")
            self.pathSaveFileKeys.append(f"{report}_Template_PDF_Location")

        #settings for auto report generation
        checkingFrequencyLabelsLayout=QGridLayout()
        checkingFrequencyLabelsLayout=self.createCheckingFrequencyLabelsLayout()
        
        #adding file paths to layout
        excelPathLabelsLayout=QVBoxLayout()
        excelPathLabelsLayout.addWidget(self.reportPathHeading)
        for ExcelFilePathKeys in self.pathSaveFileKeys:
            if "_Excel_Path" in ExcelFilePathKeys:
                excelPathLabelsLayout.addLayout(self.createLabelTextPair(f"{ExcelFilePathKeys}"))
                
        self.templatePDFLocationHeading=QLabel("Template PDF Locations")
        templatePdfLabelsLayout=QVBoxLayout()
        templatePdfLabelsLayout.addWidget(self.templatePDFLocationHeading)
        for TemplatePDFLocationKeys in self.pathSaveFileKeys:
            if "_Template_PDF_Location" in TemplatePDFLocationKeys:
                templatePdfLabelsLayout.addLayout(self.createLabelTextPair(f"{TemplatePDFLocationKeys}"))

        #combining all different layouts
        self.settingsLayout.addLayout(checkingFrequencyLabelsLayout, 1, 1)
        self.settingsLayout.addLayout(excelPathLabelsLayout, 2, 1)
        self.settingsLayout.addLayout(templatePdfLabelsLayout, 3, 1)

        #creating dummy qwidget object and replacing with scrollbar enabled widget
        self.scrollBar=QScrollArea()
        tempWidget=QWidget()
        tempWidget.setLayout(self.settingsLayout)
        self.scrollBar.setWidgetResizable(True)
        self.scrollBar.setWidget(tempWidget)

        #Save button not included in scroll area
        self.saveButton=QPushButton("Save")

        #adding all layout and widget to main widget
        self.mainLayout=QGridLayout()
        self.mainLayout.addWidget(self.scrollBar)
        self.mainLayout.addWidget(self.saveButton)
        self.setLayout(self.mainLayout)

        self.saveButton.clicked.connect(self.saveSettingsAction)
        self.verifyDirectories()

       
    def createLabelTextPair(self,pairName:str)->QGridLayout:
        """creating label text box and a button in a horizontal direction with background label"""

        tempLayout=QGridLayout()
        label=QLabel(pairName.replace("_", " "))
        pathTextBox=QLineEdit()
        pathTextBox.setStyleSheet("border: 1px solid black;")
        #setting object name to find it later
        pathTextBox.setObjectName(f"pathTextBox_{pairName}")
        pathTextBox.setText(settingsSaveFile.get(pairName,"Browse file location"))#loading save file and saving to textbox here
        button=QPushButton("Browse")
        button.setObjectName(f"{pairName}")
        button.clicked.connect(self.browseButtonAction)
        backgroundLabel=QLabel()
        backgroundLabel.setObjectName(f"backgroundLabel_{pairName}")
        backgroundLabel.setStyleSheet("background-color: rgba(255, 0, 0, 50);"
                           "border: 1px solid black;")
        
        tempLayout.addWidget(backgroundLabel, 1, 1, 2, 3)
        tempLayout.addWidget(label,1,1,2,1)
        tempLayout.addWidget(pathTextBox,1,2,2,1)
        tempLayout.addWidget(button,1,3,2,1)
        
        return tempLayout
    def saveSettingsAction(self):
        """saving all settings to a json file when save button clicked"""

        #creating save file if not exists
        os.makedirs(name=settingsFilePath, exist_ok=True)
        print(fr"{settingsFilePath}\{settingsfileName}")

        #taking all line edit values with object name and updating it settingsSaveFile dictionary
        for reportKey in self.pathSaveFileKeys:
            tempLineEdit=self.findChild(QLineEdit, f"pathTextBox_{reportKey}")
            if isinstance(tempLineEdit, QLineEdit):
                settingsSaveFile[reportKey]=tempLineEdit.text()
        
        #taking some other qobjects value and updating settingsSaveFile dictionary
        settingsSaveFile["Auto_check_report"]=self.autoCheckReport.isChecked()
        settingsSaveFile["Auto_generate_report"]=self.autoGenerateReport.isChecked()
        settingsSaveFile["Report_Trigger_Time"]=self.timeToTriggerReport.time().toString("hh:mm")
        settingsSaveFile["Report_Trigger_Day"]=self.dayOfWeekComboBox.currentText()
        settingsSaveFile["Weekly_Rechecking_Frequency"]=self.weeklyFrequencySpinBox.value()
        settingsSaveFile["Monthly_Rechecking_Frequency"]=self.monthlyFrequencySpinBox.value()

        #saving it as a JSON file
        with open(fr"{settingsFilePath}\{settingsfileName}", 'w') as file:
            json.dump(settingsSaveFile, file)
        #changing background label color 
        self.verifyDirectories()

        # emitting signal that settingsSaveFile dictionary has been updated
        self.settingsSavedSignal.emit()

    def browseButtonAction(self):
        """opens a file browser """
        #finding which browse button was clicked
        button=self.sender()
        if isinstance(button, QPushButton):
            key:str=button.objectName()

        #extracting path from selected file from File browser UI
        pathValue=QFileDialog.getOpenFileName(caption=f"{key}")

        #searching for lineedit of the button and updating its current value to the path
        tempPathTextBox=self.findChild(QLineEdit,f"pathTextBox_{key}")
        if isinstance(tempPathTextBox,QLineEdit) and not(pathValue[0]==""):
            tempPathTextBox.setText(pathValue[0])

        ...
    def verifyDirectories(self):
        """changes background label color """

        #looping through all keys of line edit objects extracting its value checking and updating color of background label
        for report in self.pathSaveFileKeys:
            activePath:str=settingsSaveFile.get(report,"")
            
            isDirectoryPresent=os.path.exists(activePath)
            isSupportedExtension = any(activePath.endswith(ext) for ext in supportedExcelExtensions) or activePath.endswith("pdf")
            if isDirectoryPresent is True and isSupportedExtension is True:
                label=self.findChild(QLabel, f"backgroundLabel_{report}")
                if isinstance(label, QLabel):
                    label.setStyleSheet("background-color: rgba(0, 255, 0, 50);"
                                        "border: 1px solid black;")
            else:
                label=self.findChild(QLabel, f"backgroundLabel_{report}")
                if isinstance(label, QLabel):
                    label.setStyleSheet("background-color: rgba(255, 0, 0, 50);"
                                        "border: 1px solid black;")
                    

    def createCheckingFrequencyLabelsLayout(self)->QGridLayout:
        
        """Creating Qobjects of checking frequency and aligning them in a layout"""
        autoGenerateLabelsLayout=QGridLayout()
        self.autoGenerateReport=QCheckBox("Auto generate report")
        self.autoCheckReport=QCheckBox("Auto check report")
        self.autoCheckReport.setChecked(settingsSaveFile.get("Auto_check_report", False))

        self.autoGenerateReport.setChecked(settingsSaveFile.get("Auto_generate_report", False))

        checkingFrequencyLabel=QLabel("Rechecking  Frequency")
        reportTriggerTimeLabel=QLabel("Report Trigger Time")
        self.dayOfWeekComboBox=QComboBox()
        self.dayOfWeekComboBox.addItems(["Sunday","Monday", "Tuesday", "Wednesday", "Thursday", "Friday","Saturday"])
        self.dayOfWeekComboBox.setCurrentText(settingsSaveFile.get("Report_Trigger_Day", "Monday"))
        self.timeToTriggerReport=QTimeEdit()
        self.timeToTriggerReport.setDisplayFormat("hh:mm")
        self.timeToTriggerReport.setTime(QTime.fromString(
                                                            settingsSaveFile.get("Report_Trigger_Time", 
                                                                                str(QTime.currentTime())
                                                                                                 ),"hh:mm"))
        reportTriggerDescriptionLabel=QLabel("Specify the week day and time to trigger the first check")
        weeklyReportLabel=QLabel("Weekly")
        monthlyReportLabel=QLabel("Monthly")

        self.weeklyFrequencySpinBox=QDoubleSpinBox()
        #self.weeklyFrequencySpinBox.setRange(0.5,9)
        self.weeklyFrequencySpinBox.setSingleStep(0.5)
        self.weeklyFrequencySpinBox.setValue(settingsSaveFile.get("Weekly_Rechecking_Frequency", 1))

        self.monthlyFrequencySpinBox=QDoubleSpinBox()
        #self.monthlyFrequencySpinBox.setRange(0.5,9)
        self.monthlyFrequencySpinBox.setSingleStep(0.5)
        self.monthlyFrequencySpinBox.setValue(settingsSaveFile.get("Monthly_Rechecking_Frequency", 2))
        checkingFrequencyDescription=QLabel("Specify how much hours to wait before start recheck if report is not completed")

        autoGenerateLabelsLayout.addWidget(self.autoCheckReport, 1, 1, 1, 2)
        autoGenerateLabelsLayout.addWidget(self.autoGenerateReport, 1, 2, 1, 2)
        autoGenerateLabelsLayout.addWidget(reportTriggerTimeLabel, 2, 1, 1, 4)
        autoGenerateLabelsLayout.addWidget(self.dayOfWeekComboBox, 2, 2, 1, 1)
        autoGenerateLabelsLayout.addWidget(self.timeToTriggerReport, 2, 3, 1, 1)
        autoGenerateLabelsLayout.addWidget(reportTriggerDescriptionLabel, 3, 1, 1, 4)
        autoGenerateLabelsLayout.addWidget(checkingFrequencyLabel, 4, 1, 1, 4)
        autoGenerateLabelsLayout.addWidget(weeklyReportLabel, 5, 1, 1, 1)
        autoGenerateLabelsLayout.addWidget(self.weeklyFrequencySpinBox, 5, 2, 1, 1)
        autoGenerateLabelsLayout.addWidget(monthlyReportLabel, 5, 3, 1, 1)
        autoGenerateLabelsLayout.addWidget(self.monthlyFrequencySpinBox, 5, 4, 1, 1)
        autoGenerateLabelsLayout.addWidget(checkingFrequencyDescription, 7, 1, 1, 4)
        return autoGenerateLabelsLayout
    
class individualReportLayout():
    """This class creates single report UI objects with Teams """    
    def __init__(self,reportName) -> None:
        #creating objects and setting object name for later identification
        self.layoutTitle=QLabel(reportName)
        self.reportName=reportName
        self.layoutTitle.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
        
        self.refreshButton=createButton(buttonWidth=50,buttonHeight=25,
                        imagePath=r"Icons\RefreshIcon.ico",toolTip="Refresh")
        self.refreshButton.setObjectName(f"{self.reportName}_RefreshButton")
        self.refreshButton.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
        self.buttonsList:list[QPushButton]=[]
        

        
        self.updatedTimeLabel=QLabel("Last Updated: Waiting for refresh")
        self.updatedTimeLabel.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
        self.generateReportButton=QPushButton("Generate Report")
        self.generateReportButton.setObjectName(f"{self.reportName}_GenerateReport")
        self.activeTimeofReport=QLabel("report week: ")
        self.activeTimeofReport.setObjectName(f"{self.reportName}_ActiveTime")
        
        self.reportBackgroundLabel=QLabel()
        self.reportBackgroundLabel.setAlignment(Qt.AlignmentFlag.AlignTop)

        #UI objects are arranged here
        self.Layout1=QGridLayout()
        self.Layout1.addWidget(self.reportBackgroundLabel, 1, 1,3,2)
        self.Layout1.addWidget(self.layoutTitle, 1, 1,1,3)
        self.Layout1.addWidget(self.refreshButton, 1, 2,1,1)
        self.Layout1.addWidget(self.updatedTimeLabel, 3, 1)
        self.Layout1.addWidget(self.generateReportButton, 3, 2)
        self.Layout1.addWidget(self.activeTimeofReport, 3, 3)

        

    def addButton(self,buttonWidth:int, buttonHeight:int,buttonName:str,positionX:int=0,positionY:int=0
                ,imagePath:str="",toolTip:str="",Active:bool=True,buttonDescription:str="") -> None:
        """Adds a button with an image on it """
        self.button=QPushButton(buttonName)
        self.button.setGeometry(positionX, positionY, buttonWidth, buttonHeight)
        #to resize image with button size
        self.button.setIconSize(QSize(buttonWidth, buttonHeight))
        self.button.setObjectName(f"{self.reportName}_{buttonDescription}")
        self.button.setIcon(QIcon(imagePath))
        self.button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.button.setStyleSheet("background : transparent;")
        if not(toolTip==""):
            self.button.setToolTip(toolTip)
        self.button.setToolTipDuration(2000)
        self.button.setEnabled(Active)
        self.buttonsList.append(self.button)
        # after creating buttons arranging it in the layout
        self.setDefaultLayout()

    def setDefaultLayout(self) -> None:
        """since team buttons is created dynamically layout is updated after adding buttons"""
        self.Layout1=QGridLayout()
        self.Layout1.addWidget(self.reportBackgroundLabel, 1, 1,3,len(self.buttonsList))
        self.Layout1.addWidget(self.layoutTitle, 1, 1,1,4)
        self.Layout1.addWidget(self.refreshButton, 1, len(self.buttonsList),1,1)
        for i,button in enumerate(self.buttonsList):
            self.Layout1.addWidget(button,2,i+1)
        tempLayout=QHBoxLayout()
        tempLayout.addWidget(self.updatedTimeLabel, alignment=Qt.AlignmentFlag.AlignLeft)
        tempLayout.addWidget(self.generateReportButton, stretch=1)
        tempLayout.addWidget(self.activeTimeofReport,alignment=Qt.AlignmentFlag.AlignRight)
        self.Layout1.addLayout(tempLayout, 3, 1, 1, len(self.buttonsList))


    #to get report layout instance to combine it with other report layout outside this class
    def getLayout(self):
        return self.Layout1

class consoleRedirector():
    def __init__(self,textEdit:QTextEdit) :
        self.textEdit=textEdit

    def write(self,text):
        self.textEdit.append(text)
        self.textEdit.ensureCursorVisible()
    def flush(self):
        pass

class KPIMainWindow(QMainWindow):

    def __init__(self) -> None:
        super().__init__()
        self.redirectConsoleToLineEdit()
        self.setWindowTitle(appName)
        self.setWindowIcon(QIcon(appIcon))
        self.resize(mainWindowWidth,mainWindowHeight)

        self.createUI()
        self.loadUnfilledTeamsLogic()
        self.addHeaderObjects("KPI Updated status")
        self.alignObjects()
        #self.createStatusWindow()

        #settings object is created here-----------------------------------------------------------------------------------
        self.settingsWindow=settingsWidget()
        self.settingsButton.clicked.connect(self.settingsWindow.show)
        self.recheckingTime:QDateTime
        self.messageBox=QMessageBox()
        if settingsSaveFile.get("Auto_check_report", False) :
            self.initalizeRefreshData()

    def createUI(self):
        """Creating all UI for dashboard and storing it in dictionary"""
        self.reportsLayoutDict:dict[str,individualReportLayout]={}

        #creating UI for all reports here
        for report in reports:
            self.reportsLayoutDict[report]=individualReportLayout(report)

        varButtonHeight=25
        varButtonWidth=50

        #adding teams button from external data for each report
        for report,teamsInReport in reportsAndTeamsDict.items():
            for teamAbbrevation,teamDataTuple in teamsInReport.items():
                teamData=teamDataTuple[0]
                teamFlag=teamData["icon"]
                activeLayout=self.reportsLayoutDict[report]
                activeLayout.addButton(buttonWidth=varButtonWidth, buttonHeight=varButtonHeight,
                buttonName=teamAbbrevation, imagePath=teamFlag,buttonDescription=teamData["teamName"])
                teamData=None
                teamFlag=None
        for layout in self.reportsLayoutDict.values():
            for button in layout.buttonsList:
                button.clicked.connect(self.flagButtonClickedAction)
        #All refresh buttons of reports are connected to single function
        #All generate buttons of reports are connected to single function
        for layoutObject in self.reportsLayoutDict.values():
            layoutObject.refreshButton.clicked.connect(self.refreshButtonClickedAction)
            layoutObject.generateReportButton.clicked.connect(self.generateReportButtonClickedAction)

        #used to trigger the first check procedure and not used as an UI
        self.startUnfilledRecheckProcedure=QPushButton()
        self.startUnfilledRecheckProcedure.clicked.connect(self.recheckProcedure)

    def redirectConsoleToLineEdit(self):
        self.logWindow=QWidget()
        self.logWindow.setWindowTitle("Log Window")
        self.logLineEdit=QTextEdit()
        self.logLineEdit.setReadOnly(True)
        tempLayout=QHBoxLayout()
        tempLayout.addWidget(self.logLineEdit)
        self.logWindow.setLayout(tempLayout)
        self.consoleRedirector=consoleRedirector(self.logLineEdit)
        sys.stdout=self.consoleRedirector

    def loadUnfilledTeamsLogic(self):
        """creating all report verifier objects and storing it in dictionary"""
        self.reportVerifierDict:dict[str,KPIreportVerifier]={}
        #setting weekly or monthly here------------------------------------------------------------------------
        for report in reports:
            self.reportVerifierDict[report]=KPIreportVerifier(reportName=report,checkingFrequency="Monthly")
        self.reportVerifierDict["LT & Orders"].checkingFrequency="Weekly"
        
        # if external data refresh required the excel is temporary copied, values refreshed, saved as temp file and used
        self.reportVerifierDict["LT & Orders"].isExternalDataRefreshRequired=True
        self.reportVerifierDict["Claims"].isExternalDataRefreshRequired=True

        #setting macroname and location here
        self.reportVerifierDict["LT & Orders"].MacroModule="Sheet1"
        self.reportVerifierDict["LT & Orders"].macroName="PrintToPDF"
        self.reportVerifierDict["On Time Delivery"].MacroModule="Sheet1"
        self.reportVerifierDict["On Time Delivery"].macroName="PrintToPDF"
        self.reportVerifierDict["Efficiency"].MacroModule="Sheet1"
        self.reportVerifierDict["Efficiency"].macroName="PrintToPDF"
        self.reportVerifierDict["NC"].MacroModule="Sheet1"
        self.reportVerifierDict["NC"].macroName="PrintToPDF"
        self.reportVerifierDict["Claims"].MacroModule="Sheet1"
        self.reportVerifierDict["Claims"].macroName="PrintToPDF"
        self.reportVerifierDict["Technical Sales Support"].MacroModule="ThisWorkbook"
        self.reportVerifierDict["Technical Sales Support"].macroName="PrintToPDF"

        self.reportVerifierDict["LT & Orders"].reportPDFName="Lead Time Corporate"
        self.reportVerifierDict["On Time Delivery"].reportPDFName="OTD report"
        self.reportVerifierDict["Efficiency"].reportPDFName="Efficiency"
        self.reportVerifierDict["NC"].reportPDFName="NC"
        self.reportVerifierDict["Claims"].reportPDFName="Claims"
        self.reportVerifierDict["Technical Sales Support"].reportPDFName="TSS"

        # if generate report procedure is not similar to all reports then procedure can be manually written and connected to separate instance
        self.reportVerifierDict["LT & Orders"].generateReport=types.MethodType(LTGenerateReportOverride,self.reportVerifierDict["LT & Orders"])

        #adding teams in reportverifier object
        for report,teamsInReport in reportsAndTeamsDict.items():
            for teamDataTuple in teamsInReport.values():
                teamData=teamDataTuple[0]
                self.reportVerifierDict[report].add_team(teamData["teamName"], teamDataTuple[1])
                teamData=None

    def refreshButtonClickedAction(self) -> None:
        """Checks the report , prints the unfilled teams and updates some properties of the reportverirfier object"""

        #Identifying which report's refresh button is clicked
        clickedRefreshButton=self.sender()
        isReportCheckedSuccessfully=False
        for reportKey,reportVerifier in self.reportVerifierDict.items():

            #setting excel path file from settingsSaveFile here when refreshing
            if reportKey in clickedRefreshButton.objectName():
                reportVerifier.report_location=settingsSaveFile.get(f"{reportKey}_Excel_Path","Key not found")
                #calling the reportverifier method to check and update some properties
                unfilledTeamsList=reportVerifier.get_teams_with_unfilled_cells()
                isReportCheckedSuccessfully=reportVerifier.isSuccessfullyCheckedWithoutErrors
                #updating the status of the UI of refreshed report
                tempLabel=self.findChild(QLabel, f"{reportKey}_ActiveTime")
                if reportVerifier.checkingFrequency=="Weekly" and isinstance(tempLabel,QLabel):
                    tempLabel.setText(f"Report Week: {reportWeek}")
                elif reportVerifier.checkingFrequency=="Monthly"and isinstance(tempLabel,QLabel):
                    tempLabel.setText(f"Report Month: {calendar.month_name[reportMonth]}")
                tempLabel=None
                break
        #updating the status of the UI of refreshed report
        for teamButton in self.reportsLayoutDict[reportKey].buttonsList:
            teamButton.setStyleSheet("")
        if isReportCheckedSuccessfully==True and unfilledTeamsList!=[]:
            for teamName in unfilledTeamsList:
                print(f"    {reportKey} pending {teamName}")
                for teamButton in self.reportsLayoutDict[reportKey].buttonsList:
                    if (teamName in teamButton.objectName()):
                        teamButton.setStyleSheet("background-color: rgba(255, 0, 0, 0.2)")
                    if teamButton.styleSheet()!="background-color: rgba(255, 0, 0, 0.2)":
                        teamButton.setStyleSheet("background-color: rgba(0, 255, 0, 0.2);")
            self.reportsLayoutDict[reportKey].updatedTimeLabel.setText(f"Last Updated : {formattedCurrentDatetime()}")

        elif isReportCheckedSuccessfully==True and unfilledTeamsList==[]:
            for teamButton in self.reportsLayoutDict[reportKey].buttonsList:
                teamButton.setStyleSheet("background-color: rgba(0, 255, 0, 0.2);")
                self.reportsLayoutDict[reportKey].updatedTimeLabel.setText(f"Last Updated : {formattedCurrentDatetime()}")
            print("    All teams has filled the data")

        elif isReportCheckedSuccessfully==False:
            for teamButton in self.reportsLayoutDict[reportKey].buttonsList:
                teamButton.setStyleSheet("background-color: rgba(0, 0, 0, 0);")
            self.reportsLayoutDict[reportKey].updatedTimeLabel.setText(f"Error report not updated")

    def initalizeRefreshData(self)->None:
        columns=list(dataBaseColumnsAndDataTypes.keys())
        logDatabase=Database(dataBasePath=fr"{logFilePath}")
        currentDateTime=QDateTime.currentDateTime()
        for report in reports:
            recheckDateTime=str(logDatabase.getLatestData(tableName=report,columnName=columns[8]))
            recheckDateTime=QDateTime.fromString(recheckDateTime,"dd-MM-yyyy HH:mm:ss")
            if recheckDateTime.isValid() and currentDateTime.secsTo(recheckDateTime)<0 :
                self.reportsLayoutDict[report].refreshButton.click()
            else:
                self.loadLastRefreshData(report=report)

    def loadLastRefreshData(self,report) -> None:
        logDatabase=Database(dataBasePath=fr"{logFilePath}")
        print(f"Loading Last saved data for {report}")
        reportStatus=logDatabase.getLatestRow(report)
        unFilledTeamsList:list=ast.literal_eval(reportStatus[4])
        reportButtons=self.reportsLayoutDict[report].buttonsList
        for teamButton in reportButtons:
            for teamName in unFilledTeamsList:
                if teamName in teamButton.objectName():
                    teamButton.setStyleSheet("background-color: rgba(255, 0, 0, 0.2)")
            
        for teamButton in reportButtons:
            if teamButton.styleSheet()!="background-color: rgba(255, 0, 0, 0.2)":
                teamButton.setStyleSheet("background-color: rgba(0, 255, 0, 0.2);")
        self.reportsLayoutDict[report].updatedTimeLabel.setText(f"Last Updated : {reportStatus[6]}")
            
    def refreshAllData(self) -> None:
        print("Refreshing all data")
        for layout in self.reportsLayoutDict.values():
            layout.refreshButton.click()
        
    def addHeaderObjects(self,buttonText:str) -> None:
        self.label=QLabel()
        self.label.setText(buttonText)
        self.settingsButton=createButton(buttonWidth=50,
                        buttonHeight=50,imagePath=settingsIcon,toolTip="Settings")
        self.emptyLabel=QPushButton()


    def alignObjects(self) -> None:
        """aligning all created ui objects and creating scroll bar"""
        masterLayout=QGridLayout()
        firstRow=QHBoxLayout()
        firstRow.addWidget(self.label,stretch=3)
        firstRow.addWidget(self.settingsButton,stretch=1)
        masterLayout.addLayout(firstRow,1,1)
        for counter,reportLayout in enumerate(self.reportsLayoutDict.values()):
            masterLayout.addLayout(reportLayout.getLayout(), counter+2, 1)

        tempWidget=QWidget()
        tempWidget.setLayout(masterLayout)
        self.scrollBar=QScrollArea()
        self.scrollBar.setWidgetResizable(True)
        self.scrollBar.setWidget(tempWidget)
        tempLayout=QGridLayout()
        tempLayout.addWidget(self.scrollBar)
        self.mainWidget=QWidget()
        self.mainWidget.setLayout(tempLayout)
        
        self.setCentralWidget(self.mainWidget)

    def createStatusWindow(self) -> None:
        """currently not used"""
        self.statusWindow=QStatusBar()
        self.setStatusBar(self.statusWindow)
        self.statusLabel=QLabel("Ready")
        self.statusLabel.setFixedHeight(50)
        self.statusWindow.addWidget(self.statusLabel)
        self.setWindowTitle("DashBoard")
    
    def flagButtonClickedAction(self) -> None:
        """currently not used"""
        temp=self.sender()
        if isinstance(temp, QPushButton):
            print(temp.objectName())

    def generateReportButtonClickedAction(self):
        """Triggers the generate report procedure in the reportVerifier object"""
        
        checkCompletion=True
        if isinstance(self.sender(),QPushButton):
            button=self.sender()
        else:
            print("Button not pressed")
            return

        for reportKey,reportVerifier in self.reportVerifierDict.items():
            if reportKey in button.objectName():
                activeReportName=reportKey
                activeReport=reportVerifier
                break
        if activeReport is None:
            print("Report not generated, Couldn't find which report the clicked button belongs")
            return
        #giving warning if report is not completed
        if activeReport.isEveryoneFilled==False:
            msgbox=QMessageBox()
            msgbox.setText("Report Completion check failed do you want to generate anyways?")
            msgbox.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if msgbox.exec()==QMessageBox.StandardButton.Yes:
                checkCompletion=False
            else:
                return
        # changing properties of reportVerifier object from settingsSaveFile and starting generate report procedure
        if activeReport.isEveryoneFilled==True or checkCompletion==False:
            tempPath1:str=settingsSaveFile.get(f"{reportKey}_Excel_Path", "")
            tempPath1=os.path.dirname(tempPath1)
            tempPath1=fr"{tempPath1}\{activeReport.reportPDFName}.pdf"
            tempPath2=settingsSaveFile.get(f"{reportKey}_Template_PDF_Location", "")
            activeReport.reportPDFLocation=tempPath1
            activeReport.reportTemplatePDFLocation=tempPath2
            print(f"Currently generating {activeReportName}")
            activeReport.generateReport()

    def recheckProcedure(self) -> None:
        """Used for auto triggering report checking and generating"""
        if not self.recheckingTime.isNull():
            currentDateTimeObject=QDateTime.currentDateTime()
            recheckingIntervalMS:int=currentDateTimeObject.secsTo(self.recheckingTime)
        else:
            recheckingIntervalMS:int=settingsSaveFile.get("Weekly_Rechecking_Frequency", 1)*60*60*1000
        if settingsSaveFile.get("Auto_check_report", False) :
            print(f"Auto check initiated checking for month {calendar.month_name[reportMonth]} and week:{reportWeek} \n")
        
            #using qtimer for rechecking triggers
            self.unFilledRecheckTimer:QTimer=QTimer(self)
            self.unFilledRecheckTimer.timeout.connect(self.autoCheckAndGenerateReport)
            if recheckingIntervalMS<0:
                self.autoCheckAndGenerateReport()
            else:
                self.unFilledRecheckTimer.start(recheckingIntervalMS) 
        else:
            print("Auto check not enabled")
        

    def autoCheckAndGenerateReport(self) -> None:
            """checks the report creates a message box showing status, generates report if it is enabled"""
            weeklyRecheckingFrequencyMS:int=settingsSaveFile.get("Weekly_Rechecking_Frequency", 24)*60*60*1000
            monthlyRecheckingFrequencyMS:int=settingsSaveFile.get("Monthly_Rechecking_Frequency", 3)*60*60*1000
            self.isRecheckRequired=True
            message="_________________________________________________\n"
            self.messageBox.setWindowTitle(f"Status for month {calendar.month_name[reportMonth]} and week:{reportWeek}")
            self.messageBox.setIcon(QMessageBox.Icon.Information)
            #collecting information about reports, generating only if report is completed and auto_generate_Report is enabled
            for reportKey,reportVerifier in self.reportVerifierDict.items():
                if reportVerifier.isReportGenerated==False:
                    self.reportsLayoutDict[reportKey].refreshButton.click()

                    if len(reportVerifier.unfilled_teams)==0 and settingsSaveFile.get("Auto_generate_report", False)==True :
                        self.reportsLayoutDict[reportKey].generateReportButton.click()
                        message=message+str(f"{reportKey} report generated \n")

                    elif len(reportVerifier.unfilled_teams)==0 and settingsSaveFile.get("Auto_generate_report", False)==False:
                        self.messageBox.setStandardButtons(QMessageBox.StandardButton.Ok)
                        message=message+str(f"{reportKey} report completed \n")

                    else:

                        message=message+str(f"{reportKey} report not completed \n pending teams : {reportVerifier.unfilled_teams} \n "
                                                    f"No of times rechecked : {reportVerifier.checkingIterationsRan}\n")
                                                    
                else:
                    self.messageBox.setStandardButtons(QMessageBox.StandardButton.Ok)
                    message=message+str(f"{reportKey} report already generated \n")
                    print(f"{reportKey} report already generated")
                message=message+"_________________________________________________\n"
            self.messageBox.setText(message)
            #showing all the report stats to user for all triggers happening
            self.messageBox.show()

            #if all weekly report is completed recheck timer is changed to monthly
            weeklyReportGenerationStatus=[]
            for reportVerifier in self.reportVerifierDict.values():
                if reportVerifier.checkingFrequency=="Weekly":
                    weeklyReportGenerationStatus.append(reportVerifier.isReportGenerated)

            if all(weeklyReportGenerationStatus):
                print("All weekly reports generated")
                self.unFilledRecheckTimer.stop()
                self.unFilledRecheckTimer.start(monthlyRecheckingFrequencyMS)
            else:
                print("All weekly reports not generated")
                self.unFilledRecheckTimer.stop()
                self.unFilledRecheckTimer.start(weeklyRecheckingFrequencyMS)

            if all(reportVerifier.isReportGenerated==True for reportVerifier in self.reportVerifierDict.values()):
                self.isRecheckRequired=False

            #If all completed rechecking procedure is closed
            if self.isRecheckRequired==False:
                print("Auto check completed report closing Timer")
                self.unFilledRecheckTimer.stop()

    def closeEvent(self, event):
    # closing event to destroy timer and message box
        if hasattr(self, 'unFilledRecheckTimer'):
            self.unFilledRecheckTimer.stop()
        if hasattr(self, 'messageBox'):
            self.messageBox.close()
        event.accept()

                



def createButton(buttonWidth:int, buttonHeight:int,positionX:int=0,positionY:int=0,buttonName:str=""
                ,imagePath:str="",toolTip:str="",Active:bool=True)->QPushButton:
    """creates button, sets some properties and returns it"""
    button=QPushButton(buttonName)
    
    button.setGeometry(positionX, positionY, buttonWidth, buttonHeight)
    #to resize image with button size
    button.setIconSize(QSize(buttonWidth, buttonHeight))
    button.setIcon(QIcon(imagePath))
    button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
    button.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
    if not(toolTip==""):
        button.setToolTip(toolTip)
    button.setToolTipDuration(2000)
    button.setEnabled(Active)
    return button

def formattedCurrentDatetime():
    return datetime.now().strftime("%d-%m-%Y %H:%M:%S")

def LTGenerateReportOverride(self:KPIreportVerifier):
        print(f"{self.reportName} : Report generation procedure overridden")
        if self.isEveryoneFilled==False:
            print("Report completion check failed")
        didMacroRun=runExcelMacro(excelFilePath=self.report_location, modulename=self.MacroModule, 
                                  macroName=self.macroName,saveExcelFile=True)
        isLinksCopied=copyPasteLinksofPDF(sourcePDF=self.reportTemplatePDFLocation,destinationPDF=self.reportPDFLocation)
        if not (didMacroRun and isLinksCopied):
            print(f"Report generation failed"
                    f"Macro ran ?: {didMacroRun}"
                    f"links copied ?: {isLinksCopied}")
        if didMacroRun and isLinksCopied:
            self.isReportGenerated=True
        else:
            self.isReportGenerated=False
        self.markReportGenerationStatus()
