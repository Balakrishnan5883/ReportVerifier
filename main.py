from PySide6.QtWidgets import QApplication
import sys
from mainWindowUI import KPIMainWindow

application = QApplication(sys.argv)

mainWindow = KPIMainWindow()



mainWindow.show()
sys.exit  (application.exec())


