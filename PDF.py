# +
from RPA.PDF import PDF

pdf = PDF()

class PDF_File:

    def Get_PDFText(self, pdfPath):
        try:
            text = pdf.get_text_from_pdf(pdfPath)
            return text
        except:
            raise Exception("")

    def Get_PDFNameInvestment(self, pdfText):
        try:
            sectionA = pdfText[1]
            firstPosition = sectionA.find("Name of this Investment:")
            secondPosition = sectionA.find("2. Unique Investment Identifier (UII):")
            nameInvestment = sectionA[firstPosition+24:secondPosition]
            return nameInvestment.strip()
        except:
            raise Exception("")

    def Get_PDFUII(self, pdfText):
        try:
            sectionA = pdfText[1]
            firstPosition = sectionA.find("2. Unique Investment Identifier (UII):")
            secondPosition = sectionA.find("Section B: Investment Detail")
            UIIValue = sectionA[firstPosition+38:secondPosition]
            return UIIValue.strip()
        except:
            raise Exception("")

    def find(self, rowList, target1, target2):
        try:
            for i,lst in enumerate(rowList):
                for j,valor in enumerate(lst):
                    if valor == target1 and rowList[i][j+2] == target2:
                        return True
            return False
        except:
            raise Exception("")
