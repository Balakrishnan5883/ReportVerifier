from PySide6.QtWidgets import QApplication,QSystemTrayIcon,QMenu
from PySide6.QtGui import QIcon,QAction
import sys
from mainWindowUI import KPIMainWindow
from applicationData import (reports,reportMonth,reportWeek,columnsAndDataTypes,logFilePath,appName,appIcon,settingsIcon,
                             quitIcon,settingsSaveFile)
from LogData import Database
import calendar
from apscheduler.schedulers.background import BackgroundScheduler




class Application(QApplication):
    def __init__(self):
        super().__init__(sys.argv)
        self.setStyleSheet(open("Styles.css", "r").read())
        self.mainWindow=KPIMainWindow()
        self.initalizeSystemTray()
        self.setQuitOnLastWindowClosed(False)
        self.scheduler=BackgroundScheduler()
        self.initalizeScheduler()
        self.mainWindow.settingsWindow.settingsSavedSignal.connect(self.rescheduleScheduler)
        
        sys.exit(self.exec())

    #system tray available next to time for the application to run in background
    def initalizeSystemTray(self):
        self.systemTrayIcon=QIcon(appIcon)
        self.SystemTray=QSystemTrayIcon(parent=self, icon=self.systemTrayIcon)
        self.TrayIconMenus=QMenu()
        #creating menus for showing dashboard, settings and quit options
        self.QuitAction=QAction(text="Quit", parent=self.TrayIconMenus,icon=QIcon(quitIcon))
        self.showDashBoardAction=QAction(text="Show Dashboard", parent=self.TrayIconMenus,icon=QIcon(appIcon))
        self.settingsAction=QAction(text="Settings", parent=self.TrayIconMenus,icon=QIcon(settingsIcon))
        self.TrayIconMenus.addAction(self.showDashBoardAction)
        self.TrayIconMenus.addAction(self.settingsAction)
        self.TrayIconMenus.addAction(self.QuitAction)
        self.showDashBoardAction.triggered.connect(self.mainWindow.show)
        self.settingsAction.triggered.connect(self.mainWindow.settingsWindow.show)
        self.QuitAction.triggered.connect(self.quit)
        self.SystemTray.setContextMenu(self.TrayIconMenus)
        self.SystemTray.show()
        
    #starts the timer for first check based on the settingsSaveFile
    def initalizeScheduler(self):
            reportTriggerTime:str=settingsSaveFile.get("Report_Trigger_Time", "14:50")
            reportTriggerDay:str=settingsSaveFile.get("Report_Trigger_Day", "Monday")
            self.scheduler.add_job(self.mainWindow.startUnfilledRecheckProcedure.click,
                                    'cron',day_of_week=reportTriggerDay[:3], 
                                    hour=reportTriggerTime[:2], minute=reportTriggerTime[-2:],
                                    id="StartChecking",replace_existing=True
                                    )
            self.scheduler.start()
            #check if any of the report is not generated and triggers checking procedure
            #if program stopped before trigger time and starts after trigger a trigger is skipped
            #so whenever application starts checked is triggered if any one of the report is not completed
            if CheckIfAutoCheckNeeded():
                self.mainWindow.startUnfilledRecheckProcedure.click()

    #reschedule scheduler if settings are changed and saved
    def rescheduleScheduler(self):
        reportTriggerTime:str=settingsSaveFile.get("Report_Trigger_Time", "14:50")
        reportTriggerDay:str=settingsSaveFile.get("Report_Trigger_Day", "Monday")
        # day only accepts first 3 letters, time is passed from '10:15' string to hours =10 and min=15
        self.scheduler.add_job(self.mainWindow.startUnfilledRecheckProcedure.click,
                                    'cron',day_of_week=reportTriggerDay[:3], 
                                    hour=reportTriggerTime[:2], minute=reportTriggerTime[-2:],
                                    id="StartChecking",replace_existing=True
                                    )
        
# checks if any one of the report is not filled in the database and returns bool based on that
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


