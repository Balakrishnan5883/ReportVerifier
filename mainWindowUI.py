from PySide6.QtWidgets import (QMainWindow,QWidget,QPushButton,QStatusBar,QLabel,QLineEdit
                               ,QGridLayout,QVBoxLayout,QHBoxLayout,QSizePolicy,QBoxLayout,
                               QFileDialog,QCheckBox,QScrollArea,QMessageBox)
from PySide6.QtGui import QIcon,QPixmap
from PySide6.QtCore import QSize,QPropertyAnimation,QPoint,Qt
from CheckUnfilledTeams import TeamData,KPIreportVerifier,supportedExcelExtensions, from_json
import json
from datetime import datetime
import os
from teamDatas import team_report,reports,activeMonth,activeWeek
from zoneinfo import ZoneInfo
appName="KPI reviewer"

mainWindowWidth = 750
mainWindowHeight = 750
settingsFilePath=fr"{os.path.expanduser("~")}\Documents\{appName}"
settingsfileName="settings.json"

if os.path.exists(fr"{settingsFilePath}\{settingsfileName}"):
    with open(fr"{settingsFilePath}\{settingsfileName}", 'r') as file:
        settingsSaveFile=json.load(file)
else:
    settingsSaveFile={}

class settingsWidget(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.settingsObjectRowsCount=0
        self.setWindowTitle("Settings")
        self.resize(1200, 1200)
        self.saveFileKeys=[]
        self.reportPathHeading=QLabel("Report Excel Paths")
        self.settingsLayout=QGridLayout()
        self.settingsLayout.addWidget(self.reportPathHeading,self.settingsObjectRowsCount,1)
        for report in reports:
            self.saveFileKeys.append(f"{report}_Excel_Path")
            self.saveFileKeys.append(f"{report}_Template_PDF_Location")

        for ExcelFilePathKeys in self.saveFileKeys:
            if "_Excel_Path" in ExcelFilePathKeys:
                self.settingsLayout.addLayout(self.createLabelTextPair(f"{ExcelFilePathKeys}"),self.settingsObjectRowsCount,1)
                
        self.settingsObjectRowsCount+=1
        self.templatePDFLocationHeading=QLabel("Template PDF Locations")
        self.autoGenerateReport=QCheckBox("Auto generate report")
        self.autoGenerateReport.setChecked(settingsSaveFile.get("Auto_generate_report", False))

        
        tempLayout=QHBoxLayout()
        tempLayout.addWidget(self.templatePDFLocationHeading)
        tempLayout.addWidget(self.autoGenerateReport)
        self.settingsLayout.addLayout(tempLayout, self.settingsObjectRowsCount, 1)
        tempLayout=None

        for TemplatePDFLocationKeys in self.saveFileKeys:
            if "_Template_PDF_Location" in TemplatePDFLocationKeys:
                self.settingsLayout.addLayout(self.createLabelTextPair(f"{TemplatePDFLocationKeys}"),self.settingsObjectRowsCount,1)
        
        self.saveButton=QPushButton("Save")
        self.settingsObjectRowsCount+=1
        self.settingsLayout.addWidget(self.saveButton,self.settingsObjectRowsCount,1)
        self.setLayout(self.settingsLayout)
        self.saveButton.clicked.connect(self.saveSettingsAction)
        self.verifyDirectories()

        self.scrollBar=QScrollArea()
        self.scrollBar.setWidgetResizable(True)
        self.scrollBar.setWidget(self)

        


            
    def createLabelTextPair(self,pairName:str)->QGridLayout:
        self.tempLayout=QGridLayout()
        self.label=QLabel(pairName.replace("_", " "))
        self.pathTextBox=QLineEdit()
        self.pathTextBox.setStyleSheet("border: 1px solid black;")
        self.pathTextBox.setObjectName(f"pathTextBox_{pairName}")
        self.pathTextBox.setText(settingsSaveFile.get(pairName,"Browse file location"))#loading save file and saving to textbox here
        self.button=QPushButton("Browse")
        self.button.setObjectName(f"{pairName}")
        self.button.clicked.connect(self.browseButtonAction)
        self.backgroundLabel=QLabel()
        self.backgroundLabel.setObjectName(f"backgroundLabel_{pairName}")
        self.backgroundLabel.setStyleSheet("background-color: rgba(255, 0, 0, 50);"
                           "border: 1px solid black;")
        
        self.tempLayout.addWidget(self.backgroundLabel, 1, 1, 2, 3)
        self.tempLayout.addWidget(self.label,1,1,2,1)
        self.tempLayout.addWidget(self.pathTextBox,1,2,2,1)
        self.tempLayout.addWidget(self.button,1,3,2,1)
        self.settingsObjectRowsCount+=1
        return self.tempLayout
    def saveSettingsAction(self):
        
        os.makedirs(name=settingsFilePath, exist_ok=True)
        print(fr"{settingsFilePath}\{settingsfileName}")
        for reportKey in self.saveFileKeys:
            tempLineEdit=self.findChild(QLineEdit, f"pathTextBox_{reportKey}")
            if isinstance(tempLineEdit, QLineEdit):
                settingsSaveFile[reportKey]=tempLineEdit.text()
        settingsSaveFile["Auto_generate_report"]=self.autoGenerateReport.isChecked()
        with open(fr"{settingsFilePath}\{settingsfileName}", 'w') as file:
            json.dump(settingsSaveFile, file)
        
        self.verifyDirectories()

    def browseButtonAction(self):
        self.button=self.sender()
        if isinstance(self.button, QPushButton):
            self.key=self.button.objectName()
        pathValue=QFileDialog.getOpenFileName(caption=f"{self.key}")
        self.tempPathTextBox=self.findChild(QLineEdit,f"pathTextBox_{self.key}")
        if isinstance(self.tempPathTextBox,QLineEdit) and not(pathValue[0]==""):
            self.tempPathTextBox.setText(pathValue[0])

        ...
    def verifyDirectories(self):
        for number,report in enumerate(self.saveFileKeys):
            activePath:str=settingsSaveFile.get(f"{report}","")
            isDirectoryPresent=os.path.exists(activePath)
            isSupportedExtension = any(activePath.endswith(ext) for ext in supportedExcelExtensions)
            if isDirectoryPresent is True: #and isSupportedExtension is True:
                label=self.findChild(QLabel, f"backgroundLabel_{report}")
                if isinstance(label, QLabel):
                    label.setStyleSheet("background-color: rgba(0, 255, 0, 50);"
                                        "border: 1px solid black;")
            else:
                label=self.findChild(QLabel, f"backgroundLabel_{report}")
                if isinstance(label, QLabel):
                    label.setStyleSheet("background-color: rgba(255, 0, 0, 50);"
                                        "border: 1px solid black;")
                

class individualReportLayout():
        
    def __init__(self,reportName) -> None:

        flagButtonHeight=25
        flagButtonWidth=50
        self.layoutTitle=QLabel(reportName)
        self.reportName=reportName
        self.layoutTitle.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
        self.refreshButton=createButton(buttonWidth=50,buttonHeight=25,
                        imagePath=r"Icons\RefreshIcon.jpg",toolTip="Refresh")
        self.refreshButton.setObjectName(f"{self.reportName}_RefreshButton")
        self.refreshButton.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
        self.buttonsList:list[QPushButton]=[]

        
        self.updatedTimeLabel=QLabel("Last Updated:")
        self.updatedTimeLabel.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
        self.generateReportButton=QPushButton("Generate Report")
        self.generateReportButton.setObjectName(f"{self.reportName}_GenerateReport")
        self.activeTimeofReport=QLabel("report week: ")
        self.activeTimeofReport.setObjectName(f"{self.reportName}_ActiveTime")
        
        self.reportBackgroundLabel=QLabel()
        self.reportBackgroundLabel.setAlignment(Qt.AlignmentFlag.AlignTop)

        self.masterLayout=QGridLayout()

        self.Layout1=QGridLayout()
        self.Layout1.addWidget(self.reportBackgroundLabel, 1, 1,3,2)
        self.Layout1.addWidget(self.layoutTitle, 1, 1,1,3)
        self.Layout1.addWidget(self.refreshButton, 1, 2,1,1)
        self.Layout1.addWidget(self.updatedTimeLabel, 3, 1)
        self.Layout1.addWidget(self.generateReportButton, 3, 2)
        self.Layout1.addWidget(self.activeTimeofReport, 3, 3)

        self.firstRowLayout=QHBoxLayout()
        self.firstRowLayout.addWidget(self.refreshButton)
        self.firstRowLayout.addWidget(self.layoutTitle)
        

    def addButton(self,buttonWidth:int, buttonHeight:int,buttonName:str,positionX:int=0,positionY:int=0
                ,imagePath:str="",toolTip:str="",Active:bool=True,buttonDescription:str="") -> None:
        
        self.button=QPushButton(buttonName)
        self.button.setGeometry(positionX, positionY, buttonWidth, buttonHeight)
        #to resize image with button size
        self.button.setIconSize(QSize(buttonWidth, buttonHeight))
        self.button.setObjectName(f"{self.reportName}_{buttonDescription}")
        self.button.setIcon(QIcon(imagePath))
        self.button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.button.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
        if not(toolTip==""):
            self.button.setToolTip(toolTip)
        self.button.setToolTipDuration(2000)
        self.button.setEnabled(Active)
        self.buttonsList.append(self.button)
        self.setDefaultLayout()


    def setDefaultLayout(self) -> None:
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



    def getLayout(self):
        return self.Layout1

    def disableButton(self,button:QPushButton):
        button.setEnabled(False)


class KPIMainWindow(QMainWindow):

    
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("KPI Remainder")
        
        self.reportsLayoutDict:dict[str,individualReportLayout]={}
        for report in reports:
            self.reportsLayoutDict[report]=individualReportLayout(report)
        

        self.addTeamButtons()
        self.loadUnfilledTeamsLogic()
        self.resize(mainWindowWidth,mainWindowHeight)
        self.addKPIObjects("KPI Updated status")
        self.alignObjects()
        self.settingsWindow=settingsWidget()
        self.settingsButton.clicked.connect(lambda:self.settingsWindow.scrollBar.show())
        
        for layoutName,layoutObject in self.reportsLayoutDict.items():
            layoutObject.refreshButton.clicked.connect( self.refreshButtonClickedAction)
            layoutObject.generateReportButton.clicked.connect(self.generateReport)

        self.refreshAllData()
    def addTeamButtons(self) -> None:
        varButtonHeight=25
        varButtonWidth=50
        
        for report,teamsInReport in team_report.items():
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

    def loadUnfilledTeamsLogic(self):

        self.reportList=list(settingsSaveFile.values())
        self.reportVerifierDict:dict[str,KPIreportVerifier]={}
        #setting checking frequency here------------------------------------------------------------------------
        for report in reports:

            self.reportVerifierDict[report]=KPIreportVerifier(checkingFrequency="Monthly")

        self.reportVerifierDict["LT & Orders"].checkingFrequency="Weekly"

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

        for report,teamsInReport in team_report.items():
            for teamDataTuple in teamsInReport.values():
                teamData=teamDataTuple[0]
                self.reportVerifierDict[report].add_team(teamData["teamName"], teamDataTuple[1])
                teamData=None

    def refreshButtonClickedAction(self) -> None:
        clickedRefreshButton=self.sender()
        isReportChecked=False
        for reportKey,reportVerifier in self.reportVerifierDict.items():
            if reportKey in clickedRefreshButton.objectName():
                reportVerifier.report_location=settingsSaveFile.get(f"{reportKey}_Excel_Path","Key not found")
                unfilledTeamsList=reportVerifier.get_teams_with_unfilled_cells()
                isReportChecked=reportVerifier.isReportChecked
                tempLabel=self.findChild(QLabel, f"{reportKey}_ActiveTime")
                if reportVerifier.checkingFrequency=="Weekly" and isinstance(tempLabel,QLabel):
                    tempLabel.setText(f"Report Week: {activeWeek}")
                elif reportVerifier.checkingFrequency=="Monthly"and isinstance(tempLabel,QLabel):
                    tempLabel.setText(f"Report Month: {datetime(2000,activeMonth,1).strftime('%B')}")
                tempLabel=None
                break
        for teamButton in self.reportsLayoutDict[reportKey].buttonsList:
            teamButton.setStyleSheet("")
        print(f"Currently checking {reportKey}")
        if isReportChecked==True and unfilledTeamsList!=[]:
            for teamName in unfilledTeamsList:
                print(f"    {reportKey} pending {teamName}")
                for teamButton in self.reportsLayoutDict[reportKey].buttonsList:
                    if (teamName in teamButton.objectName()):
                        teamButton.setStyleSheet("background-color: rgba(255, 0, 0, 0.2)")
                    if teamButton.styleSheet()!="background-color: rgba(255, 0, 0, 0.2)":
                        teamButton.setStyleSheet("background-color: rgba(0, 255, 0, 0.2);")
            self.reportsLayoutDict[reportKey].updatedTimeLabel.setText(f"Last Updated : {formattedCurrentDatetime()}")

        elif isReportChecked==True and unfilledTeamsList==[]:
            for teamButton in self.reportsLayoutDict[reportKey].buttonsList:
                teamButton.setStyleSheet("background-color: rgba(0, 255, 0, 0.2);")
                self.reportsLayoutDict[reportKey].updatedTimeLabel.setText(f"Last Updated : {formattedCurrentDatetime()}")
            print("    All teams has filled the data")

        elif isReportChecked==False:
            for teamButton in self.reportsLayoutDict[reportKey].buttonsList:
                teamButton.setStyleSheet("background-color: rgba(0, 0, 0, 0);")
            self.reportsLayoutDict[reportKey].updatedTimeLabel.setText(f"Error report not updated")
            
                    


    def refreshAllData(self) -> None:
        print("Refreshing all data")
        for layout in self.reportsLayoutDict.values():
            layout.refreshButton.click()
        
    def addKPIObjects(self,buttonText:str) -> None:
        self.label=QLabel()
        self.label.setText(buttonText)
        self.settingsButton=createButton(buttonWidth=50,
                        buttonHeight=50,imagePath=r"Icons\settings.png",toolTip="Settings")
        self.emptyLabel=QPushButton()


    def alignObjects(self) -> None:
        masterLayout=QGridLayout()
        firstRow=QHBoxLayout()
        firstRow.addWidget(self.label,stretch=3)
        firstRow.addWidget(self.settingsButton,stretch=1)
        masterLayout.addLayout(firstRow,1,1)
        for counter,reportLayout in enumerate(self.reportsLayoutDict.values()):
            masterLayout.addLayout(reportLayout.getLayout(), counter+2, 1)

        self.mainWidget=QWidget() 
        self.mainWidget.setLayout(masterLayout)
        self.setCentralWidget(self.mainWidget)

        self.setWindowTitle("DashBoard")
        
        
    def settingButtonClicked(self) -> None:
        print("Settings button clicked")
    
    def flagButtonClickedAction(self) -> None:
        temp=self.sender()
        if isinstance(temp, QPushButton):
            print(temp.objectName())

    def generateReport(self,checkCompletion:bool=True):

        temp=self.sender()
        if isinstance(temp, QPushButton):
            button=temp
            temp=None
        else:
            print("no Button found")
            return
        if checkCompletion==False:
            msgbox=QMessageBox()
            msgbox.setText("Report Completion check failed do you want to generate anyways?")
            msgbox.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if msgbox.exec()==QMessageBox.StandardButton.Yes:
                checkCompletion=True
        if checkCompletion==True:
            
            for reportKey,reportVerifier in self.reportVerifierDict.items():
                if reportKey in button.objectName():
                    tempPath1:str=settingsSaveFile.get(f"{reportKey}_Excel_Path", "")
                    tempPath1=tempPath1.replace(".xlsm",".pdf")
                    tempPath2=settingsSaveFile.get(f"{reportKey}_Template_PDF_Location", "")
                    reportVerifier.reportPDFLocation=tempPath1
                    reportVerifier.reportTemplatePDFLocation=tempPath2
                    reportVerifier.runExcelMacro()
                    if checkCompletion==True:reportVerifier.isReportChecked=True
                    print(f"Currently generating {reportKey}")
                    reportVerifier.generatePDFReport()
                    

        




def createButton(buttonWidth:int, buttonHeight:int,positionX:int=0,positionY:int=0,buttonName:str=""
                ,imagePath:str="",toolTip:str="",Active:bool=True)->QPushButton:
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
    return datetime.now().strftime("%d-%b-%y %I:%M %p")



