# +
from RPA.Browser.Selenium import Selenium
import time
from Excel_Class import *
from RPA.Excel.Files import Files
from Files import File_System
from RPA.FileSystem import FileSystem


libFile = FileSystem()
browser_lib = Selenium()
fileFile = File_System()
dirPath = fileFile.Get_OutputDirectory("output")
browser_lib.set_download_directory(directory= dirPath, download_pdf=False)

class Browser:

    def open_the_website(self, url):
        try:            
            browser_lib.open_available_browser(url)
            browser_lib.maximize_browser_window()
            print("Open browser.")
        except:
            raise Exception("Can't open the browser.")

    def click_DiveIn(self):
        try:
            browser_lib.click_element_if_visible("//a[@aria-controls='home-dive-in']")
        except:
            raise Exception("Can't clicn in DiveIn")

    def Get_Agencies(self):
        agencieList = []
        try:
            browser_lib.wait_until_element_is_visible("//span[@class='h4 w200']")
            agencies = browser_lib.find_elements(locator="//span[@class='h4 w200']")

            for agencie in agencies:
                agencieList.append(agencie.text)
            agencieList = list(dict.fromkeys(agencieList))
            agencieList = list(filter(None, agencieList))

            return agencieList
        except:
            raise Exception("Can't get all Agencies in the web.")

    def Get_Cost(self, agencieListQuantity):
        agencieCostValue = []
        try:
            cantidadAgencias = len(agencieListQuantity)
            agencieCost = browser_lib.find_elements(locator="//span[@class=' h1 w900']")
            del agencieCost[cantidadAgencias:]

            for agencie in agencieCost:
                agencieCostValue.append(agencie.text)

            return agencieCostValue
        except:
            raise Exception("Can't get all the cost of the Agencies")

    def Select_Agency(self, agencyToSelect):
        try:
            browser_lib.click_element_if_visible("(//span[@class='h4 w200'][normalize-space()='" + agencyToSelect + "'])[1]")
            browser_lib.wait_until_element_is_visible("//select[@aria-controls='investments-table-object']",10,"Error")
            browser_lib.select_from_list_by_value("//select[@aria-controls='investments-table-object']", "-1")
            print("Select " + agencyToSelect + " agency in the web.")
        except:
            raise Exception("Can't select " + agencyToSelect)

    def Get_UIILink(self):
        UIIListLink = []
        try:
            UIILink = browser_lib.find_elements(locator="//td[@class='left sorting_2']//a")

            for link in UIILink:
                UIIListLink.append(link.get_attribute("href"))

            return UIIListLink
        except:
            raise Exception("Can't get all the UII links from the table.")

    def Get_Table(self):
        listaTemporal = []
        listaFinal = []        
        try:
            UIIValue = browser_lib.find_elements(locator="//tr[@role='row']//td[@class='left sorting_2']")
            BureauValue = browser_lib.find_elements(locator="//tr[@role='row']//td[@class=' left select-filter']")
            InvestmentTitleValue = browser_lib.find_elements(locator="//tr[@role='row']//td[@class=' left']")
            TotalFYValue = browser_lib.find_elements(locator="//tr[@role='row']//td[@class=' right']")
            TypeValue = browser_lib.find_elements(locator="//tr[@role='row']//td[@class=' left select-filter']")
            CIOProjectValue = browser_lib.find_elements(locator="//tr[@role='row']//td[@class=' center']")

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
            print("Get Investments table.")
            return listaFinal
        except:
            raise Exception("Can't extract all the data of the table.")

    def Get_UII(self):
        try:
            AllUII = browser_lib.find_elements(locator="//td[@class='left sorting_2']")
            UIILink = browser_lib.find_elements(locator="//td[@class='left sorting_2']//a")
            UIIList = []

            for UII in AllUII:
                UIIList.append(UII.text)
            return UIIList, UIILink
        except:
            raise Exception("")

    def Download_PDF(self, urlList, path):
        try:
            for url in urlList:
                browser_lib.go_to(url)
                browser_lib.click_element_when_visible("//a[normalize-space()='Download Business Case PDF']")
                try:
                    browser_lib.wait_until_page_does_not_contain_element("//img[@alt='Generating PDF']",60,"Can't download PDF.")
                except:
                    exist = libFile.does_file_exist(path)
                    if exist == False:
                        browser_lib.go_to(url)
                        browser_lib.click_element_when_visible("//a[normalize-space()='Download Business Case PDF']")
                        try:
                            browser_lib.wait_until_page_does_not_contain_element("//img[@alt='Generating PDF']",60,"Can't download PDF.")
                        except:
                            continue
            print("Download PDFs.")
        except:
            raise Exception("Error during PDF downloads.")
# -


