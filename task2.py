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

#Variables
outputDirName = "output"

if __name__ == "__main__":
    configPath = file_System.Get_OutputDirectory("ConfigFile")    
    configValues = excel.Get_ConfigFileValue(configPath + "/ConfigFile.xlsx")
    agencyToSelect = configValues["Agency"]
    excelName = configValues["ExcelName"]      
    outputDirPath = file_System.Create_OutputDirectory(outputDirName)   
    excelPath = outputDirPath + excelName
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
    browser.Download_PDF(UIILinkList, downloadPath)
    time.sleep(5)
    print("Search PDF data in Investments table.")    
    pdf_Path = file_System.Get_PathFiles(downloadPath + "*.pdf")
    for pdf in pdf_Path:
        text = pdf_File.Get_PDFText(pdf)
        UIIValue = pdf_File.Get_PDFUII(text)
        NameInvestment = pdf_File.Get_PDFNameInvestment(text)       
        result = pdf_File.find(tableInvestments, UIIValue, NameInvestment)
        print("UIIValue = " + UIIValue + " and NameInvestment = " + NameInvestment + ". Exist " + result)


# -



