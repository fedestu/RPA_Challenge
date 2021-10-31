# +
from Browser_Class import Browser
from Files import File_System
from PDF import PDF_File
from Excel_Class import ExcelAgency
import time

browser = Browser()
file_System = File_System()
pdf_File = PDF_File()
excel = ExcelAgency()

# +
if name == "main":
    excel.Create_Woorkbook("C:/Users/Programacion/Desktop/Prueba.xlsx")
    browser.open_the_website("http://itdashboard.gov/")
    browser.click_DiveIn()
    AgencyList = browser.Get_Agencies()
    AgencyCost = browser.Get_Cost(AgencyList)
    excel.Write_AgencyData(AgencyList, AgencyCost)
    browser.Select_Agency("Department of Commerce")
    time.sleep(12)
    tableInvestments = browser.Get_Table()
    UIILinkList = browser.Get_UIILink()
    browser.Download_PDF(UIILinkList)
    excel.Create_NewWorksheet("Individual Investments")
    excel.Write_TableHeaders()
    excel.Write_IITable(tableInvestments)
    pdf_Path = file_System.Get_PathFiles("C:/Users/Programacion/Desktop/Output/*.pdf")
    for pdf in pdf_Path:
        pdf_File.Get_PDFText(pdf)
        pdf


# -


