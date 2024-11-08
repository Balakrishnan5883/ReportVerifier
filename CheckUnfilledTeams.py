import datetime
import os
import win32com.client as win32
from datetime import datetime
import calendar


import openpyxl
import openpyxl.cell
import openpyxl.workbook
import openpyxl.worksheet
import openpyxl.worksheet.worksheet

from copyPasteLinksOfPDF import copyPasteLinksofPDF
from LogData import Database
from teamDatas import columnsAndDataTypes,logFilePath,reportMonth,reportWeek


supportedExcelExtensions=[".xlsx",".xlsm",".xltx",".xltm"]


class KPIreportVerifier:
    def __init__(self,reportName:str,checkingFrequency:str="",reportPDFLocation:str=""
                 ,reportTemplatePDFLocation:str=""):
        self.report_location:str
        self.teams = []
        self.unfilled_teams = []
        self.responsibleData:dict[str,dict[str,list[str]]]={}
        self.reportName:str=reportName
        self.isEveryoneFilled=False
        self.isExcelPathPresent=False
        self.isReportGenerated=False
        self.isSuccessfullyCheckedWithoutErrors=False
        self.checkingFrequency=checkingFrequency
        self.reportPDFLocation:str=reportPDFLocation
        self.reportTemplatePDFLocation:str=reportTemplatePDFLocation
        self.MacroModule:str
        self.macroName:str
        self.isExternalDataRefreshRequired:bool=False
        self.tempReportPath:str=""
        self.checkingIterationsRan:int
        
        self.getReportGeneratedStatus()

    def refreshAndSaveInTempPath(self):
        if self.tempReportPath and os.path.exists(self.tempReportPath):
            try:
                os.remove(self.tempReportPath)
            except Exception as e:
                print(f"Error removing temp file: {self.tempReportPath}\n{e}")
        excel = win32.DispatchEx("Excel.Application")
        excel.DisplayAlerts=False
        excel.AskToUpdateLinks = False
        workbook = excel.Workbooks.Open(Filename=self.report_location,UpdateLinks=True)
        excel.CalculateUntilAsyncQueriesDone()
        tempFileName = f"temp_{self.reportName}_{datetime.now().strftime('%d-%b-%y_%I-%M_%p')}.xlsm"
        self.tempReportPath ="\\".join(self.report_location.split("/")[:-1])+"\\"+tempFileName
        try:
            workbook.SaveAs(Filename=self.tempReportPath, FileFormat=52)  
        except Exception as error:
            print(f"Error while saving {self.tempReportPath}\n{error}")
        finally:
            workbook.Close(SaveChanges=True)
            excel.Quit()
            workbook=  None
            excel = None
        

    def add_team(self, team_data,responsibleData:dict[str,list[str]]):
        self.teams.append(team_data)
        self.responsibleData[team_data]=responsibleData

    def getResponsibleData(self,team_data):
        return self.responsibleData[team_data]
    
    def getResponsibleSheets(self, team_data):
        return list(self.responsibleData[team_data].keys())
    
    def getResponsibleCells(self, team_data):
        return self.responsibleData[team_data]

    def get_teams_with_unfilled_cells(self):
        self.unfilled_teams:list[str]=[]
        print(f"Currently checking {self.reportName}")
        print (f"No of times Rechecked : {self.checkingIterationsRan} ")
        self.isTeamAreDefined = bool(len(self.teams))
        self.isResponsibleDataPresent = bool(len(self.responsibleData))
        self.isExcelPathPresent=False
        if os.path.exists(self.report_location):
            self.isExcelPathPresent = any(self.report_location.endswith(ext) for ext in supportedExcelExtensions)

        if not (self.isTeamAreDefined and self.isResponsibleDataPresent and self.isExcelPathPresent ):
            print(f"    Error While checking team\n"
                  f"    TeamsDeclared:{self.isTeamAreDefined}\n"
                  f"    ResponsibleDataDeclared:{self.isResponsibleDataPresent}\n"
                  f"    ExcelPathDeclared:{self.isExcelPathPresent}\n")
            return list(self.teams)

        if self.isExternalDataRefreshRequired:
            self.refreshAndSaveInTempPath()
            currentReportPath=self.tempReportPath
        else:
            currentReportPath=self.report_location

        workbook = openpyxl.load_workbook(currentReportPath,read_only=True,data_only=True)
        for team_data, sheet_cell_dict in self.responsibleData.items():
            for sheet_name, cell_list in sheet_cell_dict.items():
                if sheet_name in workbook.sheetnames:
                    worksheet:openpyxl.worksheet.worksheet.Worksheet = workbook[sheet_name]
                else:
                    print(f"Sheet {sheet_name} not found in the workbook.")
                    return list(self.teams)
                for cell_address in cell_list:
                    try:
                        cell_value: list[openpyxl.cell.Cell] = worksheet[cell_address].value 
                    except AttributeError as error:
                        self.unfilled_teams.append(team_data)
                        print(f"{team_data} Cell address:{cell_address} is not found in {sheet_name}")
                        break
                    if not (isinstance(cell_value,int) or isinstance(cell_value,float)):
                        self.unfilled_teams.append(team_data)
                        break
        workbook.close()
        if len(self.unfilled_teams)==0:
            self.isEveryoneFilled=True
        self.logToDatabase()
        self.isSuccessfullyCheckedWithoutErrors=True
        if self.isExternalDataRefreshRequired:
            os.remove(self.tempReportPath)
            self.tempReportPath=""
        return list(set(self.unfilled_teams))
    
    def getReportGeneratedStatus(self)->bool:
        column=list(columnsAndDataTypes.keys())
        logDatabase=Database(dataBasePath=fr"{logFilePath}",TableName=self.reportName,columnsAndDataTypes=columnsAndDataTypes)
        logDatabase.cursor.execute(f"""SELECT * FROM '{self.reportName}' 
                                   WHERE {column[1]} = ? AND
                                        {column[2]} = ?
                                    ORDER BY {column[7]} DESC
                                    LIMIT 1""",(calendar.month_name[reportMonth],reportWeek))
        row=logDatabase.cursor.fetchone()
        if not(row==None or len(row)==0):
            if row[5]=='True':
                self.isReportGenerated:bool=True
            self.checkingIterationsRan=int(row[7])
        else:
            self.isReportGenerated:bool=False
            self.checkingIterationsRan=0
            

        logDatabase.disconnect(saveChanges=True)
        return self.isReportGenerated
    
    def logToDatabase(self):  
        logData:dict={}
        column=list(columnsAndDataTypes.keys()) 
        logData[column[1]]=calendar.month_name[reportMonth]
        logData[column[2]]=reportWeek
        logData[column[3]]=str(self.isEveryoneFilled)
        logData[column[4]]=str(self.unfilled_teams)
        logData[column[5]]=str(self.isReportGenerated)
        logData[column[6]]=str(datetime.now())
        
        logDatabase=Database(dataBasePath=fr"{logFilePath}",TableName=self.reportName,columnsAndDataTypes=columnsAndDataTypes)
        logDatabase.cursor.execute(f"""SELECT * FROM '{self.reportName}' 
                                   WHERE {column[1]} = ? AND
                                        {column[2]} = ?
                                    ORDER BY {column[7]} DESC
                                    LIMIT 1""",(calendar.month_name[reportMonth],reportWeek))
        row=logDatabase.cursor.fetchone()
        if row==None or len(row)==0:
            logData[column[7]]=1
        else:
            logData[column[7]]=int(row[7])+1
            self.checkingIterationsRan=int(row[7])+1
        logDatabase.insertData(data=logData,saveChanges=True)
        logDatabase.disconnect(saveChanges=True)

    def generateReport(self) -> None:        
        if self.isEveryoneFilled==False:
            print("Report completion check failed")
        didMacroRun=runExcelMacro(excelFilePath=self.report_location, modulename=self.MacroModule, macroName=self.macroName)
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
        
    def markReportGenerationStatus(self):
        column=list(columnsAndDataTypes.keys())
        logDatabase=Database(dataBasePath=fr"{logFilePath}",TableName=self.reportName,columnsAndDataTypes=columnsAndDataTypes)
        logDatabase.cursor.execute(f"""SELECT * FROM '{self.reportName}' 
                                   WHERE {column[1]} = ? AND
                                        {column[2]} = ?
                                    ORDER BY {column[7]} DESC
                                    LIMIT 1""",(calendar.month_name[reportMonth],reportWeek))
        row=logDatabase.cursor.fetchone()
        if not(row==None or len(row)==0):
            rowID=row[0]
            logDatabase.cursor.execute(f"""UPDATE '{self.reportName}'
                                           SET {column[5]} = ?
                                           WHERE {column[0]} = ?
                                        """,(str(self.isReportGenerated),rowID))

        logDatabase.disconnect(saveChanges=True)

def readExcelCell(excelFilePath: str, sheetName: str, cellAddress: str):
    if not os.path.exists(excelFilePath):
        print("Excel file not found")
        return None
        
    workbook = openpyxl.load_workbook(excelFilePath, read_only=True, data_only=True)
    if sheetName not in workbook.sheetnames:
        print(f"Sheet {sheetName} not found in the workbook.")
        return None
    worksheet = workbook[sheetName]
    try:
        cell_value = worksheet[cellAddress].value
    except AttributeError as error:
        print(f"Cell address:{cellAddress} is not found in {sheetName}")
        return None
    finally:
        if 'workbook' in locals():
            workbook.close()
        worksheet = None
        workbook = None
    return cell_value

        
def runExcelMacro(excelFilePath:str,modulename:str,macroName:str)-> bool:
        if os.path.exists(excelFilePath)==False:
            print("Excel file not found")
            return False
        if (modulename=="" or macroName==""):
            print("Macro not defined")
            return False
        excelApp=win32.DispatchEx("Excel.Application")
        excelApp.Visible=False
        workBook=excelApp.Workbooks.Open(Filename=excelFilePath,UpdateLinks=True)
        workBook.RefreshAll()
        excelApp.CalculateUntilAsyncQueriesDone()
        fileName=excelFilePath.split(r"/")[-1]
        try:
            excelApp.Application.Run(f"'{fileName}'!{modulename}.{macroName}")
            print(f"successfully ran macro {macroName}")
        except Exception as e:
            print(f"Error running macro: {e}")
            return False
        finally:
            workBook.Close(SaveChanges=True)
            excelApp.Quit()
            workBook = None
            excelApp = None
            return True
            
        
    

if __name__ == "__main__":
    ...    

