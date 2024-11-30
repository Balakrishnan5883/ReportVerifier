from PySide6.QtWidgets import QApplication,QSystemTrayIcon,QMenu
from PySide6.QtGui import QIcon,QAction
from PySide6.QtCore import QDateTime
import sys
from mainWindowUI import KPIMainWindow
from typing import Optional
from applicationData import (reports,reportMonth,reportWeek,dataBaseColumnsAndDataTypes,logFilePath,appName,appIcon,settingsIcon,
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
        self.showLogWindow=QAction(text="Show Log Window", parent=self.TrayIconMenus, icon=QIcon(appIcon))

        self.TrayIconMenus.addAction(self.showDashBoardAction)
        self.TrayIconMenus.addAction(self.settingsAction)
        self.TrayIconMenus.addAction(self.showLogWindow)
        self.TrayIconMenus.addAction(self.QuitAction)

        self.showDashBoardAction.triggered.connect(self.mainWindow.show)
        self.settingsAction.triggered.connect(self.mainWindow.settingsWindow.show)
        self.showLogWindow.triggered.connect(self.mainWindow.logWindow.show)
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
            recheckingTime=CheckIfAutoCheckNeeded()
            if not recheckingTime is None:
                self.mainWindow.recheckingTime=recheckingTime
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
        
# checks if any one of the report is not filled in the database 
def CheckIfAutoCheckNeeded() -> Optional[QDateTime]:
    completionDictionary:dict[str,bool]={}
    recheckingTimeList:list[QDateTime]=[]
    columns=list(dataBaseColumnsAndDataTypes.keys())
    logDatabase=Database(dataBasePath=fr"{logFilePath}")
    for report in reports:
        latestData=logDatabase.getLatestRow(tableName=report)
        reportMonthName=str(calendar.month_name[reportMonth])
        if reportMonthName in latestData and str(reportWeek) in latestData:
            completionDictionary[report]=logDatabase.getLatestData(tableName=report, columnName=columns[5])=='True'
            tempRecheckDateTime:str=str(logDatabase.getLatestData(tableName=report,columnName=columns[8]))
            recheckDateTime:QDateTime=QDateTime.fromString(tempRecheckDateTime, "dd-MM-yyyy HH:mm:ss")
            if recheckDateTime.isValid():
                recheckingTimeList.append(recheckDateTime)
        else:
            print(f"CheckIfAutoCheckNeeded failed cannot find the report for month {reportMonth} and week {reportWeek}")
            return QDateTime.currentDateTime()
    logDatabase.disconnect(saveChanges=True)

    if any(completion is False for completion in completionDictionary.values()): 
        if len(recheckingTimeList)==0:
            return QDateTime.currentDateTime()
        else:
            return getMinimumQDateTime(recheckingTimeList)
    else:
        return None

def getMinimumQDateTime(listOfDateTime:list[QDateTime])-> Optional[QDateTime]:
    if len(listOfDateTime)==0:
        return None
    else:
        minimumDateTime=listOfDateTime[0]
        for dateTime in listOfDateTime[1:0]:
            if dateTime.secsTo(minimumDateTime)<0:
                minimumDateTime=dateTime
        return minimumDateTime

if __name__=="__main__":
    application=Application()


