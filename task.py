from ITDashboard import IT_Dashboard

if __name__ == "__main__":
    itDashboard = IT_Dashboard()
    try:
        print("Initializing the process.")
        print("Obtaining data from the agencies.")
        itDashboard.get_Agencies_Cost()
        print("Writing the information (name and cost) of the agencies in the excel.")
        itDashboard.write_Agencies_Cost()
        print("Obtaining the table data of the selected agency from the ConfigFile.")
        agencyTable = itDashboard.get_Agency_Table()
        print("Writing the data from the agency table select in excel.")
        itDashboard.write_Agency_Table(agencyTable)
        print("Downloading the corresponding PDFs if available. Please be patient.")
        itDashboard.download_PDFs()        
        print("Check that the data in the downloaded PDFs match the data extracted from the table.")
        itDashboard.find_Exist_Value(agencyTable)
        print("The process has been successfully completed. ")
    finally:
        itDashboard.close_Browser()

