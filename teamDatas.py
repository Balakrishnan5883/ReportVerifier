from datetime import datetime
import os,sys

WIT:dict[str,str] = {"teamName":"Wittur Italy","icon":"Icons\\ItalyFlag.jpeg", "teamLeader":"John","abbrevation":"WIT"}
WES:dict[str,str] = {"teamName":"Wittur Spain","icon":"Icons\\SpainFlag.png", "teamLeader":"Bill","abbrevation":"WES Doors"}
WHU:dict[str,str] = {"teamName":"Wittur Hungary","icon":"Icons\\HungaryFlag.jpg", "teamLeader":"Arpad","abbrevation":"WHU"}
SSC:dict[str,str] = {"teamName":"Shared Service Center","icon":"Icons\\IndiaFlag.png", "teamLeader":"Lance","abbrevation":"SSC"}
WAT:dict[str,str] = {"teamName":"Wittur Austria","icon":"Icons\\AustriaFlag.png", "teamLeader":"Dan","abbrevation":"WAT Slings"}
WAR:dict[str,str] = {"teamName":"Wittur Argentina","icon":"Icons\\ArgentinaFlag.jpg", "teamLeader":"Dave","abbrevation":"WAR"}

reportWeek:int = datetime.now().isocalendar()[1]
reportMonth:int = datetime.now().month-1


LTActiveRowIndex=reportWeek+1
OTDActiveRowIndex=reportWeek+1
def getColumnAlphabetfromNumber(column_number:int)->str:
    result = ""
    while column_number > 0:
        column_number -= 1
        remainder = column_number % 26
        result = chr(65 + remainder) + result
        column_number //= 26
    return result

NCActiveColumnIndex=getColumnAlphabetfromNumber(reportMonth+4)
ClaimsActiveColumnIndex=getColumnAlphabetfromNumber(reportMonth+4)
TSSActiveRowIndex=reportMonth+1


reportsAndTeamsDict:dict[
                        str,dict
                                [str,tuple
                                        [dict,dict]
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
        SSC['abbrevation']:(SSC,{'Sheet1':[f'H{OTDActiveRowIndex}']}),
        WAR['abbrevation']:(WAR,{'Sheet1':[f'G{OTDActiveRowIndex}']}),    
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
# changing the order of column requires update on here, on CheckUnfilledTeams.KPIreportVerifier.logToDatabase
columnsAndDataTypes:dict[str,str]={"id" :"INTEGER PRIMARY KEY",
                                "reportMonth": "TEXT"
                               ,"reportWeek":"TEXT"
                               ,"isEveryoneFilled":"TEXT"
                                ,"unfilledTeams":"TEXT"
                                ,"isReportGenerated":"TEXT"
                                ,"reportCheckedTime":"TEXT"
                                ,"iterationsRan":"INTEGER"}

reports=list(reportsAndTeamsDict.keys())
workingFolder:str=os.path.dirname(os.path.abspath(sys.argv[0]))
logFilePath:str=fr"{workingFolder}\log\log.db"