import sqlite3
import os
import sys
from typing import Optional,Union
from functools import wraps

class Database:
    def __init__(self,dataBasePath:str,TableName:str="",columnsAndDataTypes:dict[str,str]={}) -> None:
        os.makedirs(os.path.dirname(dataBasePath), exist_ok=True) 
        self.tableAndColumnsDict:dict[str, list[str]]={}
        self.connection=sqlite3.connect(f"{dataBasePath}")
        self.cursor=self.connection.cursor()
        self.tablesList:list[str]=[]
        if TableName=="" or len(columnsAndDataTypes.keys())==0:
            self.connection.commit()
        else:
            self.createTable(TableName, columnsAndDataTypes)
            self.tableAndColumnsDict[TableName]=list(columnsAndDataTypes.keys())
        self.getTables()

    @staticmethod
    def _checkTableExists(method):
        @wraps(method)
        def wrapper(self, tableName:str,  *args, **kwargs):
            if tableName not in self.tablesList:
                print(f"Table {tableName} not found")
                return None
            return method(self,tableName ,*args, **kwargs)
        return wrapper
    
    @staticmethod
    def _checkColumnExists(method):
        @wraps(method)
        def wrapper(self, tableName:str, data:dict[str,str]={}, columnName:str= "" ,*args, **kwargs):
            columnsInTable=self.getColumns(tableName)
            if len(data.keys())>0:
                columnsInInput=list(data.keys())
                for column in columnsInInput:
                    if column not in columnsInTable:
                        print(f"Column {column} not found in the table {tableName}")
                        return
                    else:
                        return method(self, tableName, data, *args, **kwargs)
            else:
                if columnName not in columnsInTable:
                    print(f"Column {columnName} not found in the table {tableName}")
                    return
                else:
                    return method(self, tableName, columnName, *args, **kwargs)
            
        return wrapper

    def getTables(self)->list[str]:
        self.cursor.execute("SELECT name FROM sqlite_master where type='table';")
        tables=self.cursor.fetchall()
        self.tablesList=list(table[0] for table in tables)
        return self.tablesList

    @_checkTableExists
    def getColumns(self, tableName:str)->list[str]:
        self.cursor.execute(f"PRAGMA table_info('{tableName}')")
        columns=self.cursor.fetchall()
        columnsList=list(column[1] for column in columns)
        return columnsList[1:]

    def createTable(self,tableName:str, columnsAndDataTypes:dict[str, str]) -> None:
        columnAndDataTypesString=",".join([f"{column} {dataType}" for column, dataType in columnsAndDataTypes.items()])
        self.cursor.execute(f"CREATE TABLE IF NOT EXISTS '{tableName}' ({columnAndDataTypesString})")    

    @_checkTableExists
    def insertColumn(self, tableName:str, columnName:str, columnDataType:str) -> None:
        self.cursor.execute(f"ALTER TABLE '{tableName}' ADD COLUMN '{columnName}' {columnDataType}")

    @_checkColumnExists
    def insertData(self, tableName,data:dict[str, str],saveChanges) -> None:
        columnsList=self.getColumns(tableName)
        columnsTuple=tuple(data.keys())
        for column in columnsTuple:
            if column not in columnsList:
                print(f"Column {column} not found in the table '{tableName}'")
                return
        valueTuple=tuple(data.values())
        self.cursor.execute(f"""INSERT INTO '{tableName}' {columnsTuple} VALUES{valueTuple}""")
        if saveChanges==True:
            self.connection.commit()
    
    @_checkTableExists
    def getAllData(self,tableName:str)->list[tuple[str]] :
        self.cursor.execute(f"SELECT * FROM '{tableName}'")
        rows=self.cursor.fetchall()
        return rows
    
    @_checkTableExists
    def clearAllData(self,tableName:str):
        self.cursor.execute(f"DELETE FROM '{tableName}'")

    @_checkTableExists
    def getLatestRow(self,tableName:str)->list[str]:
        self.cursor.execute(f"SELECT * FROM '{tableName}' ORDER BY id DESC LIMIT 1")
        row=self.cursor.fetchone()
        return list(row)
    
    @_checkTableExists
    @_checkColumnExists
    def getLatestData(self, tableName:str, columnName:str)->Union[str,int]:
        self.cursor.execute(f'SELECT "{columnName}" FROM "{tableName}" ORDER BY id DESC LIMIT 1')
        row=self.cursor.fetchone()
        return row[0]
    
    @_checkTableExists
    @_checkColumnExists
    def changeLatestData(self, tableName:str, columnName:str, value:str):
        latestRowID=self.getLatestRow(tableName)[0]
        self.cursor.execute(f"""UPDATE '{tableName}'
                               SET '{columnName}' = ?
                               WHERE id = ?
                            """,(value,latestRowID))
        
    def disconnect(self,saveChanges:bool):
        if saveChanges==True:
            self.connection.commit()
        self.connection.close()

if __name__=="__main__":
        
    columnsAndDataTypes:dict[str,str]={"id" :"INTEGER PRIMARY KEY",
                                    "reportMonth": "TEXT"
                                    ,"reportWeek":"TEXT"
                                     ,"isEveryoneFilled":"TEXT"
                                    ,"unfilledTeams":"TEXT"
                                    ,"isReportGenerated":"TEXT"
                                    ,"reportCheckedTime":"TEXT"
                                    ,"iterationsRan":"INTEGER"
                                    ,"nextRecheckTime":"TEXT"}


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
    db1=Database(dataBasePath=fr"{scriptFolder}\log\test.db")
    
    print (db1.tablesList)

    print(db1.getLatestRow(tableName="TestTable"))
    db1.changeLatestData(tableName="Testable",columnName="iterationsRan", value=str(105))
    print(db1.getLatestData("TestTable",columnName="iterationsRan"))
    print(db1.getColumns("TestTable"))
    db1.disconnect(saveChanges=True)
    

