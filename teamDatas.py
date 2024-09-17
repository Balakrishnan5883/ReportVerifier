from datetime import datetime

WIT:dict[str,str] = {"teamName":"Wittur Italy","icon":"Icons\\ItalyFlag.jpeg", "teamLeader":"John","abbrevation":"WIT"}
WES:dict[str,str] = {"teamName":"Wittur Spain","icon":"Icons\\SpainFlag.png", "teamLeader":"Bill","abbrevation":"WES"}
WHU:dict[str,str] = {"teamName":"Wittur Hungary","icon":"Icons\\HungaryFlag.jpg", "teamLeader":"Arpad","abbrevation":"WHU"}
SSC:dict[str,str] = {"teamName":"Shared Service Center","icon":"Icons\\IndiaFlag.png", "teamLeader":"Lance","abbrevation":"SSC"}
WAT:dict[str,str] = {"teamName":"Wittur Austria","icon":"Icons\\AustriaFlag.png", "teamLeader":"Dan","abbrevation":"WAT"}
WAR:dict[str,str] = {"teamName":"Wittur Argentina","icon":"Icons\\ArgentinaFlag.jpg", "teamLeader":"Dave","abbrevation":"WAR"}

activeWeek:int = datetime.now().isocalendar()[1]
activeMonth:int = datetime.now().month


LTActiveRowIndex=activeWeek+1
OTDActiveRowIndex=activeMonth+1


team_report:dict[str,dict] = {
    "LT & Orders": 
    {
        'WIT':(WIT, {'Sheet1':[f'C{LTActiveRowIndex}']}),
        'WES':(WES,{'Sheet1':[f'D{LTActiveRowIndex}']}),
        'WHU':(WHU,{'Sheet1':[f'F{LTActiveRowIndex}']}),
        'SSC':(SSC,{'Sheet1':[f'H{LTActiveRowIndex}']}),
        'WAT':(WAT,{'Sheet1':[f'E{LTActiveRowIndex}']}),
        'WAR':(WAR,{'Sheet1':[f'G{LTActiveRowIndex}']}),

    },
    "On Time Delivery": 
    {
        'WIT':(WIT, {'Sheet1':[f'C{OTDActiveRowIndex}']}),
        'WES':(WES,{'Sheet1':[f'D{OTDActiveRowIndex}']}),
        'SSC':(SSC,{'Sheet1':[f'H{OTDActiveRowIndex}']}),
        'WAR':(WAR,{'Sheet1':[f'G{OTDActiveRowIndex}']}),    
    },
    "Efficiency":
    {
        'WIT':(WIT, {'Sheet1':['C2']}),
        'WES':(WES, {'Sheet1':['D2']}),
        'WHU':(WHU, {'Sheet1':['F2']}),
        'SSC':(SSC, {'Sheet1':['H2']}),
        'WAT':(WAT, {'Sheet1':['E2']}),
        'WAR':(WAR, {'Sheet1':['G2']}),
    },
    "NC":
    {
        'WIT':(WIT, {'Sheet1':['C2']}),
        'WES':(WES, {'Sheet1':['D2']}),
        'SSC':(SSC, {'Sheet1':['H2']}),
        'WAT':(WAT, {'Sheet1':['E2']}),
        'WHU':(WES, {'Sheet1':['F2']}),
        'WAR':(WES, {'Sheet1':['G2']}),
    },
    "Claims":
    {
        'WIT':(WIT, {'Sheet1':['C2']}),
        'WES':(WES, {'Sheet1':['D2']}),
        'SSC':(SSC, {'Sheet1':['H2']}),
        'WAT':(WAT, {'Sheet1':['E2']}),
        'WAR':(WAR, {'Sheet1':['G2']}),
        'WHU':(WHU, {'Sheet1':['F2']}),

    },
    "Technical Sales Support":
    {
        'WIT':(WIT, {'Sheet1':['C2']}),
        'WES':(WES, {'Sheet1':['D2']}),
        'SSC':(SSC, {'Sheet1':['H2']}),
        'WAT':(WAT, {'Sheet1':['E2']}),
        'WAR':(WAR, {'Sheet1':['G2']}),
        'WHU':(WHU, {'Sheet1':['F2']}),


    }
}

reports=list(team_report.keys())
