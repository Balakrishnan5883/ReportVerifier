import win32com.client as win32
import os

excelApp=win32.gencache.EnsureDispatch('Excel.Application')
filePath=r'C:\Users\Bala krishnan\OneDrive\Documents\Python projects\KPI Application\Test Data\LT\Lead Time Corporate.xlsm'
workbook=excelApp.Workbooks.Open(filePath)
workbook.RefreshAll()
tempFileName="te1mp.xlsm"
tempFilePath=os.path.dirname(filePath)+f"\\{tempFileName}"
workbook.SaveAs(Filename=tempFilePath,FileFormat=52)
workbook.Close(SaveChanges=0)