# +
from RPA.FileSystem import FileSystem
from PDF import *

lib = FileSystem()

class File_System:

    def Move_Files(self, filesToMove, destinationPath):
        try:
            filesList = lib.find_files(filesToMove)
            if filesList:
                lib.move_files(filesList, destinationPath)
            return
        except:
            raise Exception("Can't move file from " + filesToMove + " to " + destinationPath)

    def Get_PathFiles(self, filesPath):
        try:
            filesList = lib.find_files(filesPath)
            return filesList
        except:
            raise Exception("Error while obtaining the path of the pdf files.")
