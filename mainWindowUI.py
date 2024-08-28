from PySide6.QtWidgets import (QMainWindow,QWidget,QPushButton,QStatusBar,QLabel
                               ,QGridLayout,QVBoxLayout,QHBoxLayout,QSizePolicy)
from PySide6.QtGui import QIcon,QPixmap
from PySide6.QtCore import QSize,QPropertyAnimation,QPoint,Qt

mainWindowWidth = 750
mainWindowHeight = 750

class individualReportLayout():
        
    def __init__(self,reportName) -> None:

        flagButtonHeight=25
        flagButtonWidth=50
        self.layoutTitle=QLabel(reportName)
        self.layoutTitle.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
        self.leadTimeRefreshButton=createButton(buttonWidth=50,buttonHeight=25,
                        imagePath=r"Icons\RefreshIcon.jpg",toolTip="Refresh")
        self.leadTimeRefreshButton.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
        self.buttonsList=[]

        self.WITButton=createButton(buttonWidth=flagButtonWidth,buttonName="WIT",
                        buttonHeight=flagButtonHeight,imagePath=r"Icons\ItalyFlagGreyed.jpeg",toolTip="Italy",Active=False)
        self.WATButton=createButton(buttonWidth=flagButtonWidth,buttonName="WAT",
                        buttonHeight=flagButtonHeight,imagePath=r"Icons\AustriaFlagGreyed.png",toolTip="Austria",Active=False)
        self.SSCButton=createButton(buttonWidth=flagButtonWidth,buttonName="SSC",
                        buttonHeight=flagButtonHeight,imagePath=r"Icons\IndiaFlagGreyed.png",toolTip="India",Active=False)
        self.WHUButton=createButton(buttonWidth=flagButtonWidth,buttonName="WHU",
                        buttonHeight=flagButtonHeight,imagePath=r"Icons\HungaryFlagGreyed.jpg",toolTip="Hungary",Active=False)
        self.WESButton=createButton(buttonWidth=flagButtonWidth,buttonName="WES",
                        buttonHeight=flagButtonHeight,imagePath=r"Icons\SpainFlagGreyed.png",toolTip="Spain",Active=False)
        
        self.updateStatusLabel=QLabel("Last Updated:")
        self.updateStatusLabel.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
        
        self.flagBackgroundLabel=QLabel()
        self.flagBackgroundLabel.setAlignment(Qt.AlignmentFlag.AlignTop)
        #self.flagBackgroundLabel.setMinimumHeight(10)
        self.flagBackgroundLabel.setStyleSheet(
                "background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 darkgrey, stop:1 black);"
                "padding: 25px;"
                )
        self.masterLayout=QGridLayout

        self.Layout1=QGridLayout()
        self.Layout1.addWidget(self.flagBackgroundLabel, 1, 1,3,5)
        self.Layout1.addWidget(self.layoutTitle, 1, 1,1,4)
        self.Layout1.addWidget(self.leadTimeRefreshButton, 1, 5,1,1)
        self.Layout1.addWidget(self.WITButton,2,1)
        self.Layout1.addWidget(self.WATButton,2,2)
        self.Layout1.addWidget(self.SSCButton,2,3)
        self.Layout1.addWidget(self.WHUButton,2,4)
        self.Layout1.addWidget(self.WESButton,2,5)
        self.Layout1.addWidget(self.updateStatusLabel, 3, 1, 1, 5)

        self.firstRowLayout=QHBoxLayout()
        self.firstRowLayout.addWidget(self.leadTimeRefreshButton)
        self.firstRowLayout.addWidget(self.layoutTitle)

    def addButton(self,buttonWidth:int, buttonHeight:int,positionX:int=0,positionY:int=0,buttonName:str=""
                ,imagePath:str="",toolTip:str="",Active:bool=True) -> None:
        
        self.button=QPushButton(buttonName)
        self.button.setGeometry(positionX, positionY, buttonWidth, buttonHeight)
        #to resize image with button size
        self.button.setIconSize(QSize(buttonWidth, buttonHeight))
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
        self.Layout1.addWidget(self.flagBackgroundLabel, 1, 1,3,len(self.buttonsList))
        self.Layout1.addWidget(self.layoutTitle, 1, 1,1,4)
        self.Layout1.addWidget(self.leadTimeRefreshButton, 1, len(self.buttonsList),1,1)
        for i,button in enumerate(self.buttonsList):
            self.Layout1.addWidget(button,2,i+1)
        self.Layout1.addWidget(self.updateStatusLabel, 3, 1, 1, 5)

        

    def getLayout(self):
        return self.Layout1

    def disableButton(self,button:QPushButton):
        button.setEnabled(False)


class KPIMainWindow(QWidget):

    
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("KPI Remainder")


        self.LTlayout=individualReportLayout("Lead Time & Orders")
        self.OTDlayout=individualReportLayout("On Time Delivery")
        self.Efficiencylayout=individualReportLayout("Efficiency")
        self.NClayout=individualReportLayout("NonConformance")
        self.ClaimsLayout=individualReportLayout("Claims")
        self.TSSLayout=individualReportLayout("Technical Sales Support")

        self.resize(mainWindowWidth,mainWindowHeight)
        self.addKPIObjects("KPI Updated status")
        self.alignObjects()


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
        masterLayout.addLayout(self.LTlayout.getLayout(), 2, 1)
        masterLayout.addLayout(self.OTDlayout.getLayout(), 3, 1)
        masterLayout.addLayout(self.Efficiencylayout.getLayout(), 4, 1)
        masterLayout.addLayout(self.NClayout.getLayout(), 5, 1)
        masterLayout.addLayout(self.ClaimsLayout.getLayout(), 6, 1)
        masterLayout.addLayout(self.TSSLayout.getLayout(), 7, 1)
        
        self.setLayout(masterLayout)
        self.setWindowTitle("DashBoard")
        
        
    def settingButtonClicked(self) -> None:
        print("Settings button clicked")

def buttonPushAnimation(button:QPushButton):
    print("Settings button clicked")
    animation=QPropertyAnimation(button,b"pos")
    animation.setEndValue(QPoint(button.x(),button.y()+250))
    animation.setDuration(2500)
    animation.start()

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



