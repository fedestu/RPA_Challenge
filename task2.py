# +
from Browser_Class import Browser
from Files import File_System
from PDF import PDF_File
from Excel_Class import ExcelAgency
import time
from RPA.Desktop.OperatingSystem import OperatingSystem

browser = Browser()
file_System = File_System()
pdf_File = PDF_File()
excel = ExcelAgency()

#Variables
outputDirName = "Output"
agencyToSelect = "National Archives and Records Administration"

if __name__ == "__main__":
    outputDirPath = file_System.Create_OutputDirectory(outputDirName)
    excelPath = outputDirPath + "Agencies.xlsx"
    downloadPath = outputDirPath + "/"    
    excel.Create_Woorkbook(excelPath)    
    browser.open_the_website("http://itdashboard.gov/")
    browser.click_DiveIn()
    AgencyList = browser.Get_Agencies()
    AgencyCost = browser.Get_Cost(AgencyList)
    excel.Write_AgencyData(AgencyList, AgencyCost, excelPath)    
    browser.Select_Agency(agencyToSelect)
    time.sleep(12)    
    tableInvestments = browser.Get_Table()    
    excel.Create_NewWorksheet("Individual Investments", excelPath)
    excel.Write_TableHeaders(excelPath)
    excel.Write_IITable(tableInvestments, excelPath)    
    UIILinkList = browser.Get_UIILink()
    browser.Download_PDF(UIILinkList)
    time.sleep(5)    
    pdf_Path = file_System.Get_PathFiles(downloadPath + "*.pdf")
    for pdf in pdf_Path:
        text = pdf_File.Get_PDFText(pdf)
        UIIValue = pdf_File.Get_PDFUII(text)
        NameInvestment = pdf_File.Get_PDFNameInvestment(text)
        print("UIIValue = " + UIIValue + " y NameInvestment = " + NameInvestment)
        result = pdf_File.find(tableInvestments, UIIValue, NameInvestment)
        print(result)


# -



