from datetime import datetime
import os,sys,json
from typing import Union
#all individual team related data is stored as key and value pair
WIT:dict[str,str] = {"teamName":"Wittur Italy","icon":"Icons\\ItalyFlag.ico", "teamLeader":"John","abbrevation":"WIT"}
WES:dict[str,str] = {"teamName":"Wittur Spain","icon":"Icons\\SpainFlag.ico", "teamLeader":"Bill","abbrevation":"WES Doors"}
WHU:dict[str,str] = {"teamName":"Wittur Hungary","icon":"Icons\\HungaryFlag.ico", "teamLeader":"Arpad","abbrevation":"WHU"}
SSC:dict[str,str] = {"teamName":"Shared Service Center","icon":"Icons\\IndiaFlag.ico", "teamLeader":"Lance","abbrevation":"SSC"}
WAT:dict[str,str] = {"teamName":"Wittur Austria","icon":"Icons\\AustriaFlag.ico", "teamLeader":"Dan","abbrevation":"WAT Slings"}
WAR:dict[str,str] = {"teamName":"Wittur Argentina","icon":"Icons\\ArgentinaFlag.ico", "teamLeader":"Dave","abbrevation":"WAR"}

#some of application settings
appName="KPI reviewer"
appIcon=r"Icons\appIcon.ico"
settingsIcon=r"Icons\settings.ico"
quitIcon=r"Icons\quit.ico"

mainWindowWidth = 750
mainWindowHeight = 750
settingsFilePath=fr"{os.path.expanduser("~")}\Documents\{appName}"
settingsfileName="settings.json"

#Loading save file stored locally in documents if found 
if os.path.exists(fr"{settingsFilePath}\{settingsfileName}"):
    with open(fr"{settingsFilePath}\{settingsfileName}", 'r') as file:
        settingsSaveFile=json.load(file)
else:
    settingsSaveFile={}

#week and month in integer 
reportWeek:int = datetime.now().isocalendar()[1]
reportMonth:int = datetime.now().month-1

#relative position of cell row or column varies based on reports and (month or week)
LTActiveRowIndex=reportWeek+1
OTDActiveRowIndex=reportMonth+1
def getColumnAlphabetfromNumber(column_number:int)->str:
    result = ""
    while column_number > 0:
        column_number -= 1
        remainder = column_number % 26
        result = chr(65 + remainder) + result
        column_number //= 26
    return result

NCActiveColumnIndex=getColumnAlphabetfromNumber(reportMonth+3)
ClaimsActiveColumnIndex=getColumnAlphabetfromNumber(reportMonth+3)
TSSActiveRowIndex=reportMonth+1


#The main data of nested dictionary combinning reports and teams with their responsible cells of the active week or month
#UI objects and reportVerifier object are created with this data
reportsAndTeamsDict:dict[
                        str,dict
                                [str,tuple
                                        [dict[str,str],
                                         dict[str,list[str]]]
                                ]
                        ] = {
    "LT & Orders": 
    {
        WIT['abbrevation']:(WIT, {'Sheet1':[f'C{LTActiveRowIndex}']}),
        WES['abbrevation']:(WES,{'Sheet1':[f'D{LTActiveRowIndex}']}),
        WHU['abbrevation']:(WHU,{'Sheet1':[f'F{LTActiveRowIndex}']}),
        SSC['abbrevation']:(SSC,{'Sheet1':[f'H{LTActiveRowIndex}']}),
        WAT['abbrevation']:(WAT,{'Sheet1':[f'E{LTActiveRowIndex}']}),
        WAR['abbrevation']:(WAR,{'Sheet1':[f'G{LTActiveRowIndex}']}),

    },
    "On Time Delivery": 
    {
        WIT['abbrevation']:(WIT, {'Sheet1':[f'C{OTDActiveRowIndex}']}),
        WES['abbrevation']:(WES,{'Sheet1':[f'D{OTDActiveRowIndex}']}),
        SSC['abbrevation']:(SSC,{'Sheet1':[f'F{OTDActiveRowIndex}']}),
        WAR['abbrevation']:(WAR,{'Sheet1':[f'E{OTDActiveRowIndex}']}),    
    },
    "Efficiency":
    {
        WIT['abbrevation']:(WIT, {'Sheet1':['C2']}),
        WES['abbrevation']:(WES, {'Sheet1':['D2']}),
        WHU['abbrevation']:(WHU, {'Sheet1':['F2']}),
        SSC['abbrevation']:(SSC, {'Sheet1':['H2']}),
        WAT['abbrevation']:(WAT, {'Sheet1':['E2']}),
        WAR['abbrevation']:(WAR, {'Sheet1':['G2']}),
    },
    "NC":
    {
        WIT['abbrevation']:(WIT, {'Sheet1':[f'{NCActiveColumnIndex}5']}),
        WES['abbrevation']:(WES, {'Sheet1':[f'{NCActiveColumnIndex}6']}),
        SSC['abbrevation']:(SSC, {'Sheet1':[f'{NCActiveColumnIndex}10']}),
        WAT['abbrevation']:(WAT, {'Sheet1':[f'{NCActiveColumnIndex}7']}),
        WHU['abbrevation']:(WES, {'Sheet1':[f'{NCActiveColumnIndex}8']}),
        WAR['abbrevation']:(WES, {'Sheet1':[f'{NCActiveColumnIndex}9']}),
    },
    "Claims":
    {
        WIT['abbrevation']:(WIT, {'Sheet1':[f'{ClaimsActiveColumnIndex}5']}),
        WES['abbrevation']:(WES, {'Sheet1':[f'{ClaimsActiveColumnIndex}6']}),
        SSC['abbrevation']:(SSC, {'Sheet1':[f'{ClaimsActiveColumnIndex}10']}),
        WAT['abbrevation']:(WAT, {'Sheet1':[f'{ClaimsActiveColumnIndex}7']}),
        WAR['abbrevation']:(WAR, {'Sheet1':[f'{ClaimsActiveColumnIndex}9']}),
        WHU['abbrevation']:(WHU, {'Sheet1':[f'{ClaimsActiveColumnIndex}8']}),

    },
    "Technical Sales Support":
    {
        WIT['abbrevation']:(WIT, {'Sheet1':[f'C{TSSActiveRowIndex}']}),
        WES['abbrevation']:(WES, {'Sheet1':[f'D{TSSActiveRowIndex}']}),
        SSC['abbrevation']:(SSC, {'Sheet1':[f'H{TSSActiveRowIndex}']}),
        WAT['abbrevation']:(WAT, {'Sheet1':[f'E{TSSActiveRowIndex}']}),
        WAR['abbrevation']:(WAR, {'Sheet1':[f'G{TSSActiveRowIndex}']}),
        WHU['abbrevation']:(WHU, {'Sheet1':[f'F{TSSActiveRowIndex}']}),


    }
}
#columns and datatype stored in database for auto generate report purposes
# changing the order of column requires update on here, on CheckUnfilledTeams.KPIreportVerifier.logToDatabase
dataBaseColumnsAndDataTypes:dict[str,str]={"id" :"INTEGER PRIMARY KEY",
                                "reportMonth": "TEXT"
                               ,"reportWeek":"TEXT"
                               ,"isEveryoneFilled":"TEXT"
                                ,"unfilledTeams":"TEXT"
                                ,"isReportGenerated":"TEXT"
                                ,"reportCheckedTime":"TEXT"
                                ,"iterationsRan":"INTEGER"
                                ,"nextRecheckTime":"TEXT"}



reports=list(reportsAndTeamsDict.keys())
workingFolder:str=os.path.dirname(os.path.abspath(sys.argv[0]))
#Database is stored where the program is located
logFilePath:str=fr"{workingFolder}\log\log.db"