# +
from RPA.Excel.Files import Files

lib = Files()

class ExcelAgency:

    def Create_Woorkbook(self, excelPath):
        try:
            lib.create_workbook(path=excelPath, fmt='xlsx')
            lib.rename_worksheet("Sheet", "Agencies")
            lib.set_cell_value(1, 1, "Agencie", name='Agencies', fmt=None)
            lib.set_cell_value(1, 2, "Cost", name='Agencies', fmt=None)
            lib.save_workbook()
        except:
            raise Exception("Error in excel creation.")

    def Write_AgencyData(self, agencyList, agencyCost):
        try:
            lib.open_workbook("C:/Users/Programacion/Desktop/Prueba.xlsx")
            row = 2
            colunm = 1
            for agency in agencyList:
                lib.set_cell_value(row, colunm, agency, name='Agencies', fmt=None)
                row += 1

            row = 2
            colunm = 2
            for cost in agencyCost:
                lib.set_cell_value(row, colunm, cost, name='Agencies', fmt=None)
                row += 1
            lib.save_workbook()
        except:
            raise Exception("Error while writing the agency data in excel.")    

    def Write_TableHeaders(self):
        try:
            lib.open_workbook("C:/Users/Programacion/Desktop/Prueba.xlsx")    
            lib.set_cell_value(1, 1, "UII", name='Individual Investments', fmt=None)
            lib.set_cell_value(1, 2, "Bureau", name='Individual Investments', fmt=None)
            lib.set_cell_value(1, 3, "Investment Title", name='Individual Investments', fmt=None)
            lib.set_cell_value(1, 4, "Total FY2021 Spending ($M)", name='Individual Investments', fmt=None)
            lib.set_cell_value(1, 5, "Type", name='Individual Investments', fmt=None)
            lib.set_cell_value(1, 6, "CIO Rating", name='Individual Investments', fmt=None)
            lib.set_cell_value(1, 7, "# of Projects", name='Individual Investments', fmt=None)
            lib.save_workbook()
        except:
            raise Exception("Error while writing table headers in excel.")

    def Create_NewWorksheet(self, wsName):
        try:
            lib.open_workbook("C:/Users/Programacion/Desktop/Prueba.xlsx")
            lib.create_worksheet(wsName, content=None, exist_ok=False, header=False)
            lib.set_cell_value(1, 1, "UII", name='Individual Investments', fmt=None)
            lib.set_cell_value(1, 2, "Bureau", name='Individual Investments', fmt=None)
            lib.set_cell_value(1, 3, "Investment Title", name='Individual Investments', fmt=None)
            lib.set_cell_value(1, 4, "Total FY2021 Spending ($M)", name='Individual Investments', fmt=None)
            lib.set_cell_value(1, 5, "Type", name='Individual Investments', fmt=None)
            lib.set_cell_value(1, 6, "CIO Rating", name='Individual Investments', fmt=None)
            lib.set_cell_value(1, 7, "# of Projects", name='Individual Investments', fmt=None)
            lib.save_workbook()
        except:
            raise Exception("Error while writing table headers in excel.")

    def Write_IITable(self, IITable):
        try:
            lib.open_workbook("C:/Users/Programacion/Desktop/Prueba.xlsx")

            i = 0
            row = 2

            for item in IITable:
                lib.set_cell_value(row, 1, IITable[i][0], name='Individual Investments', fmt=None)
                lib.set_cell_value(row, 2, IITable[i][1], name='Individual Investments', fmt=None)
                lib.set_cell_value(row, 3, IITable[i][2], name='Individual Investments', fmt=None)
                lib.set_cell_value(row, 4, IITable[i][3], name='Individual Investments', fmt=None)
                lib.set_cell_value(row, 5, IITable[i][4], name='Individual Investments', fmt=None)
                lib.set_cell_value(row, 6, IITable[i][5], name='Individual Investments', fmt=None)
                lib.set_cell_value(row, 7, IITable[i][6], name='Individual Investments', fmt=None)
                row += 1
                i += 1

            lib.save_workbook()
        except:
            raise Exception("Error while writing the table in excel.")
