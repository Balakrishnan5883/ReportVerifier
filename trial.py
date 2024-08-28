from PySide6.QtWidgets import QApplication, QPushButton, QWidget, QVBoxLayout,QLabel,QFrame,QSizePolicy,QBoxLayout,QGridLayout
from PySide6.QtCore import QPropertyAnimation, QPoint,QSize,QObject
from PySide6.QtGui import QIcon,QPixmap,Qt,QImage

class mainApp(QWidget):
    
    def __init__(self):
        super().__init__()
        self.labelText = QLabel("Click Me")
        self.labelText.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon=QIcon(r"Icons\RefreshIcon.jpg")
        self.pixmap=QPixmap(self.icon.pixmap(QSize(500,100)))
        self.labelIcon=QLabel()
        self.labelIcon.setPixmap(self.pixmap)
        self.labelIcon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.labelBackground=QLabel()
        self.labelBackground.setStyleSheet("background-color: grey;")
        self.label1=createPicturedLabel(r"Icons\SpainFlag.png","WES")
        

        layout = QBoxLayout(QBoxLayout.Direction.LeftToRight)
        layout.addLayout(self.label1)
        layout.addWidget(self.labelBackground)
        layout.addWidget(self.labelIcon)
        layout.addWidget(self.labelText)

        self.button1=createButton(buttonWidth=100,buttonHeight=100,buttonName="Button1")
        self.button1.setObjectName("Button1 Check Ok")
        self.button1.clicked.connect(self.buttonAction)
        layout.addWidget(self.button1)


        
        self.setLayout(layout)

    def buttonAction(self):
        temp=self.sender()
        if isinstance(temp, QPushButton):
            if (temp.text() == "Button1"):
                temp.setText("Button Clicked!")
            else:
                temp.setText("Button1")
        
        
        print (temp.findChild(QObject))
        

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
    button.setAccessibleName(buttonName)
    return button

def createPicturedLabel(imagePath:str,labelName:str)->QGridLayout:
    labelBackground=QLabel()
    labelText=QLabel(labelName)
    labelText.setAlignment(Qt.AlignmentFlag.AlignCenter)
    labelText.setStyleSheet("""
        color: black;
        font-weight: bold;
        font-size: 20px;
    """)
    labelImage=QLabel()
    labelImage.setAlignment(Qt.AlignmentFlag.AlignCenter)
    icon=QIcon(imagePath)
    pixmap=QPixmap(icon.pixmap(QSize(250, 100)))
    labelImage.setPixmap(pixmap)
    labelBackground.setStyleSheet("background-color: blue;")

    layout=QGridLayout()
    #layout.addWidget(labelBackground, 1, 1,2,1)
    layout.addWidget(labelImage,1,1)
    layout.addWidget(labelText,1,1)
    return layout


if __name__ == "__main__":
    app=QApplication([])
    window = mainApp()
    window.show()
    app.exec()
