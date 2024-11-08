import sqlite3
import os
import sys

class Database:
    def __init__(self,dataBasePath:str,TableName:str="",columnsAndDataTypes:dict[str,str]={}) -> None:
        self.tableName=TableName
        os.makedirs(os.path.dirname(dataBasePath), exist_ok=True) 
        self.connection=sqlite3.connect(f"{dataBasePath}")
        self.cursor=self.connection.cursor()
        
        if TableName=="" or len(columnsAndDataTypes.keys())==0:
            return
        
        columnAndDataTypesString=",".join([f"{column} {dataType}" for column, dataType in columnsAndDataTypes.items()])
        self.cursor.execute(f"CREATE TABLE IF NOT EXISTS '{self.tableName}' ({columnAndDataTypesString})")       
        self.columnsList:list=list(columnsAndDataTypes.keys())

        
    def insertData(self, data:dict,saveChanges) -> None:
        columnsTuple=tuple(data.keys())
        for column in columnsTuple:
            if column not in self.columnsList:
                print(f"Column {column} not found in the table {self.tableName}")
                return
        valueTuple=tuple(data.values())
        self.cursor.execute(f"""INSERT INTO '{self.tableName}' {columnsTuple} VALUES{valueTuple}""")
        if saveChanges==True:
            self.connection.commit()


    def displayAllData(self) -> None:
        print(f"All Data in {self.tableName}")
        self.cursor.execute(f"SELECT * FROM {self.tableName}")
        rows=self.cursor.fetchall()
        for row in rows:
            print(row)
    def clearAllData(self):
        self.cursor.execute(f"DELETE FROM {self.tableName}")
    def getLatestRecord(self):
        print(f"Latest Record in {self.tableName}")
        self.cursor.execute(f"SELECT * FROM {self.tableName} ORDER BY reportWeek DESC,reportMonth DESC LIMIT 1")
        row=self.cursor.fetchall()
        print(row)
    
    def disconnect(self,saveChanges:bool):
        if saveChanges==True:
            self.connection.commit()
        self.connection.close()

if __name__=="__main__":
        
    columnsAndDataTypes:dict[str,str]={"reportMonth": "TEXT"
                               ,"reportWeek":"TEXT"
                               ,"isEveryoneFilled":"TEXT"
                                ,"unfilledTeams":"TEXT"
                                ,"isReportGenerated":"TEXT"
                                ,"reportCheckedTime":"TEXT"
                                ,"iterationsRan":"INTEGER"}


    data1:dict[str,str]={"reportMonth": "2","reportWeek":"3","isEveryoneFilled":"True",
                        "unfilledTeams":"WIT","isReportGenerated":"True",
                        "reportCheckedTime":"2024-01-01 12:00:00","iterationsRan":"1"}
    data2:dict[str,str]={"reportMonth": "3","reportWeek":"2","isEveryoneFilled":"True",
                        "unfilledTeams":"WIT","isReportGenerated":"True",
                        "reportCheckedTime":"2024-01-01 12:00:00","iterationsRan":"1"}
    data3:dict[str,str]={"reportMonth": "4","reportWeek":"1","isEveryoneFilled":"True",
                        "unfilledTeams":"WIT","isReportGenerated":"True",
                        "reportCheckedTime":"2024-01-01 12:00:00","iterationsRan":"1"}

    scriptFolder=os.path.dirname(os.path.abspath(sys.argv[0]))
    db1=Database(dataBasePath=fr"{scriptFolder}\log\log.db",TableName="KPIReport",columnsAndDataTypes=columnsAndDataTypes)
    db1.insertData(data=data1,saveChanges=True)
    db1.insertData(data=data2,saveChanges=True)
    db1.insertData(data=data3,saveChanges=True)

    
    db1.getLatestRecord()
    db1.displayAllData()
    db1.disconnect(saveChanges=True)
    

