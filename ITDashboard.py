from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
import time
import os
import shutil

class IT_Dashboard:

    configFile = {}
    agenciesData = {}
    UIIListLink = []
    outputFolder = ""

    def __init__(self):
        self.browser = Selenium()
        self.excel = Files()
        self.pdf_File = PDF()
        self.file_System = FileSystem()

        #If you want to enter more global data, enter them in the excel and then get their value from here
        self.excel.open_workbook(os.getcwd() + "/ConfigFile/ConfigFile.xlsx")
        self.configFile["Agency"] = self.excel.get_cell_value(1, 2, name=None)
        self.configFile["ExcelName"] = self.excel.get_cell_value(2, 2, name=None)
        self.configFile["URL"] = self.excel.get_cell_value(3, 2, name=None)
        self.configFile["OutputName"] = self.excel.get_cell_value(4, 2, name=None)
        self.excel.close_workbook()
        
        self.outputFolder = os.getcwd() + "/" + self.configFile["OutputName"] + "/"
        if not os.path.exists(self.outputFolder):
            os.mkdir(self.outputFolder) 
        else:            
            shutil.rmtree(self.outputFolder)
            os.mkdir(self.outputFolder)

        self.browser.set_download_directory(self.outputFolder)        

    def get_Agencies_Cost(self):
        agencieList = []
        agencieCostValue = []        
        try:
            #Open browser
            self.browser.open_available_browser(self.configFile["URL"])
            self.browser.maximize_browser_window()

            self.browser.click_element_if_visible("//a[@aria-controls='home-dive-in']")
            #Get agencies name and cost
            self.browser.wait_until_element_is_visible("//span[@class='h4 w200']")
            agencies = self.browser.find_elements(locator="//span[@class='h4 w200']")

            for agencie in agencies:
                agencieList.append(agencie.text)
            agencieList = list(dict.fromkeys(agencieList))
            agencieList = list(filter(None, agencieList))

            cantidadAgencias = len(agencieList)
            agencieCost = self.browser.find_elements(locator="//span[@class=' h1 w900']")
            del agencieCost[cantidadAgencias:]
            for agencie in agencieCost:
                agencieCostValue.append(agencie.text)
            self.agenciesData = {"Agencies": agencieList, "Cost": agencieCostValue}

        except:
            raise Exception("Error obtaining the name and cost of the agencies.")

    def close_Browser(self):
        try:
            self.browser.close_browser()
        except:
            raise Exception("An error occurred while closing the browser.")

    def write_Agencies_Cost(self):
        try:            
            self.excel.create_workbook(self.outputFolder
             + self.configFile["ExcelName"]).append_worksheet("Sheet", self.agenciesData, True)
            self.excel.rename_worksheet("Sheet", "Agencies")
            self.excel.save_workbook()
        except:
            raise Exception("Error while writing the name and cost of the agencies in excel.")

    def get_Agency_Table(self):
        listaTemporal = []
        listaFinal = [["UII", "Bureau", "Investment Title","Total FY2021 Spending ($M)", "Type", "CIO Rating", "# Of Projects"]]   
        try:
            #Select Agency
            self.browser.click_element_if_visible("(//span[@class='h4 w200'][normalize-space()='" + self.configFile["Agency"] + "'])[1]")
            self.browser.wait_until_element_is_visible("//select[@aria-controls='investments-table-object']", 15,"Error loading the page.")
            self.browser.select_from_list_by_value("//select[@aria-controls='investments-table-object']", "-1")
            self.browser.wait_until_element_is_visible("//a[@class='paginate_button next disabled']", 15, "Error loading table.") 
            #Get UII Links
            UIILink = self.browser.find_elements(locator="//td[@class='left sorting_2']//a")

            for link in UIILink:
                self.UIIListLink.append((link.get_attribute("href"), link.text))                        
            #Get HTML Table
            UIIValue = self.browser.find_elements(locator="//tr[@role='row']//td[@class='left sorting_2']")
            BureauValue = self.browser.find_elements(locator="//tr[@role='row']//td[@class=' left select-filter']")
            InvestmentTitleValue = self.browser.find_elements(locator="//tr[@role='row']//td[@class=' left']")
            TotalFYValue = self.browser.find_elements(locator="//tr[@role='row']//td[@class=' right']")
            TypeValue = self.browser.find_elements(locator="//tr[@role='row']//td[@class=' left select-filter']")
            CIOProjectValue = self.browser.find_elements(locator="//tr[@role='row']//td[@class=' center']")

            cantidadItems = len(UIIValue)
            i = 0
            z = 1

            while i != cantidadItems:
                listaTemporal.append((UIIValue[i].text, BureauValue[i].text, InvestmentTitleValue[i].text,
                                   TotalFYValue[i].text, TypeValue[i].text, CIOProjectValue[z-1].text,
                                   CIOProjectValue[z].text))
                i += 1
                z += 2

            for item in listaTemporal:
                listaFinal.append(item)

            return listaFinal          
        except:
            raise Exception("Error when extracting all the data from the table.")

    def write_Agency_Table(self, tableAgency):
        try:
            self.excel.open_workbook(self.outputFolder + self.configFile["ExcelName"])
            self.excel.create_worksheet("Individual Investment", tableAgency, True)
            self.excel.save_workbook()
            self.excel.close_workbook()
        except:
            raise Exception("Error while writing the table of the agencies in excel.")


    def download_PDFs(self):
        try:
            for url in self.UIIListLink:
                self.browser.go_to(url[0])
                self.browser.click_element_when_visible("//a[normalize-space()='Download Business Case PDF']")
                try:
                    self.browser.wait_until_page_does_not_contain_element("//img[@alt='Generating PDF']",60,"Can't download PDF.")
                except:
                    exist = self.file_System.does_file_exist(self.outputFolder + url[1] + ".pdf")
                    if exist == False:
                        self.browser.go_to(url[0])
                        self.browser.click_element_when_visible("//a[normalize-space()='Download Business Case PDF']")
                        try:
                            self.browser.wait_until_page_does_not_contain_element("//img[@alt='Generating PDF']",60,"Can't download PDF.")
                        except:
                            print("Error during PDF download for the link " + url[0] + ".")
                            continue
            time.sleep(5)
        except:
            raise Exception("Error during PDF downloads.")    

    def find_Exist_Value(self, tableAgency):
        findValue = False
        try:
            pdf_Path = self.file_System.find_files(self.outputFolder + "*.pdf")            
            for pdf in pdf_Path:                
                text = self.pdf_File.get_text_from_pdf(pdf)                
                pdfValues = self.Get_PDFValues(text)
                for i,lst in enumerate(tableAgency):
                    for j,valor in enumerate(lst):
                        if valor == pdfValues[1] and tableAgency[i][j+2] == pdfValues[0]:
                            findValue = True                            
                if findValue:
                    print("UIIValue = " + pdfValues[1] + " and NameInvestment = " + pdfValues[0] + ". Exist = True")
                else:
                    print("UIIValue = " + pdfValues[1] + " and NameInvestment = " + pdfValues[0] + ". Exist = False")
                findValue = False
                
        except:
            raise Exception("Error when searching for the existence of UII and Inversion within the table") 

    def Get_PDFValues(self, pdfText):
        try:
            sectionA = pdfText[1]
            firstPosition = sectionA.find("Name of this Investment:")
            secondPosition = sectionA.find("2. Unique Investment Identifier (UII):")
            thirdPosition = sectionA.find("Section B: Investment Detail")
            nameInvestment = sectionA[firstPosition+24:secondPosition]
            UIIValue = sectionA[secondPosition+38:thirdPosition]            
            return nameInvestment.strip(), UIIValue.strip()
        except:
            raise Exception("Error extracting the value of the PDF.")

# if __name__ == "__main__":
#     itDashboard = IT_Dashboard()
#     try:
#         print("Initializing the process.")
#         print("Obtaining data from the agencies.")
#         itDashboard.get_Agencies_Cost()
#         print("Writing the information (name and cost) of the agencies in the excel.")
#         itDashboard.write_Agencies_Cost()
#         print("Obtaining the table data of the selected agency from the ConfigFile.")
#         agencyTable = itDashboard.get_Agency_Table()
#         print("Writing the data from the agency table select in excel.")
#         itDashboard.write_Agency_Table(agencyTable)
#         print("Downloading the corresponding PDFs if available. ")
#         itDashboard.download_PDFs()
#         time.sleep(5)
#         print("Check that the data in the downloaded PDFs match the data extracted from the table.")
#         itDashboard.find_Exist_Value(agencyTable)
#         print("The process has been successfully completed. ")
#     finally:
#         itDashboard.close_Browser() 
