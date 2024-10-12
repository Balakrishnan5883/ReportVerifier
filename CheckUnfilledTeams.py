import openpyxl
import datetime
import os
import win32com.client


import openpyxl.cell
import openpyxl.workbook
import openpyxl.worksheet
import openpyxl.worksheet.worksheet
from copyPasteLinksOfPDF import copyPasteLinksofPDF
            
supportedExcelExtensions=[".xlsx",".xlsm",".xltx",".xltm"]

class TeamData:
    def __init__(self, team_name:str, team_leader:str):
        self.teamName = team_name
        self.teamLeader = team_leader
        self.teamIconPath = ""

def from_json(json_data: dict):
    return TeamData(json_data["team_name"], json_data["team_leader"])

class KPIreportVerifier:
    def __init__(self,checkingFrequency:str="",reportPDFLocation:str="",reportTemplatePDFLocation:str=""):
        self.report_location = ""
        self.teams = []
        self.responsibleData:dict[str,dict[str,list[str]]]={}
        self.isReportChecked=False
        self.isExcelPathPresent=False
        self.reportGenerated=False
        self.checkingFrequency=checkingFrequency
        self.reportPDFLocation:str=reportPDFLocation
        self.reportTemplatePDFLocation:str=reportTemplatePDFLocation
        self.MacroModule:str=""
        self.macroName:str=""

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
        unfilled_teams = []
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
            
        workbook = openpyxl.load_workbook(self.report_location,read_only=True,data_only=True)

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
                        unfilled_teams.append(team_data)
                        print(f"{team_data} Cell address:{cell_address} is not found in {sheet_name}")
                        break
                    if not (isinstance(cell_value,int) or isinstance(cell_value,float)):
                        unfilled_teams.append(team_data)
                        break
        workbook.close()
        self.isReportChecked=True
        return list(set(unfilled_teams))
    
    def runExcelMacro(self):
        if (self.MacroModule=="" or self.macroName==""):
            print("Macro not defined")
            return
        excelApp=win32com.client.Dispatch("Excel.Application")
        excelApp.Visible=False
        workBook=excelApp.Workbooks.Open(self.report_location)
        self.reportName=self.report_location.split(r"/")[-1]
        try:
            excelApp.Application.Run(f"'{workBook.Name}'!{self.MacroModule}.{self.macroName}")
        except Exception as e:
            print(f"Error running macro: {e}")
        finally:
            workBook.Close(SaveChanges=False)
            excelApp.Quit()
            workBook = None
            excelApp = None

    def generatePDFReport(self):
        if self.isReportChecked==False:
            print("Report completion check failed")
            return
        if not (os.path.exists(self.reportPDFLocation) or os.path.exists(self.reportTemplatePDFLocation)):
            print(f"one of the PDF is not found"
                  f"report PDF exists ? : {os.path.exists(self.reportPDFLocation)}"
                  f"template PDF exists ?: {os.path.exists(self.reportTemplatePDFLocation)}")
            return
        copyPasteLinksofPDF(sourcePDF=self.reportTemplatePDFLocation,destinationPDF=self.reportPDFLocation)
        self.reportGenerated=True
        
        

if __name__ == "__main__":
    
    lead_time_report = KPIreportVerifier()
    lead_time_report.report_location=r"C:\Users\Bala krishnan\OneDrive\Documents\Python projects\Open excel and run a macro\Book1.xlsm"
    WitturItaly = TeamData(team_name="WIT", team_leader="John")
    WitturSpain = TeamData(team_name="WES", team_leader="Bill")
    WitturIndia = TeamData(team_name="WIN", team_leader="Lance")



    todaysDate=datetime.date.today()
    currentWeekNumber=todaysDate.isocalendar()[1]

    lead_time_report.add_team(WitturItaly,{"Sheet1": [f"A2", "A3"], "Sheet2": ["A2", "A3"]})
    lead_time_report.add_team(WitturSpain,{"Sheet1": ["B2", "B3"], "Sheet2": ["B2", "B3"]})
    lead_time_report.add_team(WitturIndia,{"Sheet1": ["C2", "C3"], "Sheet2": ["C2", "C3"]})

    Unfilled_teams = lead_time_report.get_teams_with_unfilled_cells()

    if len(Unfilled_teams)>0:
        print("LT & Orders KPI Report pending teams:")
        for teamData in Unfilled_teams:
            print(f"{teamData.team_name} : {teamData.team_leader}")
    else:
        print("All teams have filled all the required cells for LT & Orders KPI Report.") 
    