# +
from RPA.FileSystem import FileSystem
from PDF import *
import os
import shutil

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

    def Create_OutputDirectory(self, directoryName):
        dirName = os.getcwd() + "/" + directoryName + "/"
        if not os.path.exists(dirName):
            os.mkdir(dirName)           
            print("Directory " , dirName ,  " Created ")
        else:
            print("Directory " , dirName ,  " already exists, deleting directory.")
            shutil.rmtree(dirName)
            os.mkdir(dirName)
            print("Directory " , dirName ,  " Created ")
        return dirName
    
    def Get_OutputDirectory(self, directoryName):
        dirName = os.getcwd() + "/" + directoryName + "/"
        return dirName
