from PySide6.QtWidgets import QApplication,QSystemTrayIcon,QMenu
from PySide6.QtGui import QIcon,QAction
import sys
from mainWindowUI import KPIMainWindow
from teamDatas import (reports,reportMonth,reportWeek,columnsAndDataTypes,logFilePath)
from LogData import Database
import calendar
from apscheduler.schedulers.background import BackgroundScheduler

class Application(QApplication):
    def __init__(self):
        super().__init__(sys.argv)
        self.systemTrayIcon=QIcon(r"Icons\ItalyFlag.jpeg")
        self.mainWindow=KPIMainWindow()
        self.SystemTray=QSystemTrayIcon(parent=self, icon=self.systemTrayIcon)
        self.TrayIconMenus=QMenu()
        self.QuitAction=QAction(text="Quit", parent=self.TrayIconMenus)
        self.showDashBoardAction=QAction(text="Show Dashboard", parent=self.TrayIconMenus)
        self.TrayIconMenus.addAction(self.showDashBoardAction)
        self.TrayIconMenus.addAction(self.QuitAction)
        self.showDashBoardAction.triggered.connect(self.mainWindow.show)            
        self.QuitAction.triggered.connect(self.quit)
        self.SystemTray.setContextMenu(self.TrayIconMenus)
        #self.setStyleSheet(open("Styles.css", "r").read())
        self.SystemTray.show()
        self.SystemTray.showMessage("KPI Reviewer", "Application started", icon=QSystemTrayIcon.MessageIcon.Information)
        self.setQuitOnLastWindowClosed(False)
        self.scheduler=BackgroundScheduler()
        self.initalizeScheduler()
        
        
        sys.exit(self.exec())

    def initalizeScheduler(self):
            self.scheduler.add_job(self.mainWindow.startUnfilledRecheckProcedure.click, 'cron',day_of_week='tue', hour=22, minute=47)
            self.scheduler.start()
            #if CheckIfTriggerNeeded():
            #    self.mainWindow.initiateAutoCheck()

        

def CheckIfAutoCheckNeeded() -> bool:
    completionDictionary={}
    column=list(columnsAndDataTypes.keys())
    logDatabase=Database(dataBasePath=fr"{logFilePath}")
    for report in reports:
        logDatabase.cursor.execute(f"""SELECT * FROM '{report}' 
                                WHERE {column[1]} = ? AND
                                        {column[2]} = ?
                                    ORDER BY {column[7]} DESC
                                    LIMIT 1""",(calendar.month_name[reportMonth],reportWeek))
        row=logDatabase.cursor.fetchone()
        if not(row==None or len(row)==0):
            if row[5]=='True':
                completionDictionary[report]=True
        else:
            completionDictionary[report]=False
    logDatabase.disconnect(saveChanges=True)
    if any(completion is False for completion in completionDictionary.values()):
        return True
    else:
        return False



if __name__=="__main__":
    application=Application()


