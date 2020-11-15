
import os
import logging

import win32com.client

logger = logging.getLogger('Sun')


class pptTrans:
    def __init__(self, filePath):
        self.filePath = filePath
        self.powerpoint = None
        self.init_powerpoint()
        self.convert_files_in_folder(filePath)
        self.quit()

    def quit(self):
        if None is not self.powerpoint:
            self.powerpoint.Quit()

    def init_powerpoint(self):
        try:
            self.powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
            self.powerpoint.Visible = 2
        except Exception as e:
            logger.error(str(e))

    def ppt_to_pptx(self, inputFileName, formatType=32):
        # https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
        outputFileName = self.getNewFileName('pdf', inputFileName)
        inputFileName = inputFileName.replace('/', '\\')
        outputFileName = outputFileName.replace('/', '\\')
        if '' == outputFileName:
            return
        if None is self.powerpoint:
            return
        powerpoint = self.powerpoint

        deck = powerpoint.Presentations.Open(inputFileName)
        """
        1: ppt
        6: rtf
        7: pps
        11: pptx
        """
        try:
            deck.SaveAs(outputFileName, 32)  # formatType = 32 for ppt to pdf
        except Exception as e:
            logger.error(str(e))
        deck.Close()

    def convert_files_in_folder(self, folder):
        files = os.listdir(folder)
        pptfiles = [f for f in files if f.endswith(".ppt")]
        for pptfile in pptfiles:
            fullpath = os.path.join(folder, pptfile)
            self.ppt_to_pptx(fullpath, fullpath)

    def getNewFileName(self, newType, filePath):
        try:
            dirPath = os.path.dirname(filePath)
            baseName = os.path.basename(filePath)
            fileName = baseName.rsplit('.', 1)[0]
            suffix = baseName.rsplit('.', 1)[1]
            if newType == suffix:
                self.errorSignal.emit('类型相同')
                return ''
            newFileName = '{dir}/{fileName}.{suffix}'.format(dir=dirPath, fileName=fileName, suffix=newType)
            if os.path.exists(newFileName):
                newFileName = '{dir}/{fileName}_new.{suffix}'.format(dir=dirPath, fileName=fileName, suffix=newType)
            return newFileName
        except Exception as e:
            self.errorSignal.emit(str(e))
            return ''


if __name__ == "__main__":
    path = r'D:\Code\python\pptTrans'
    op = pptTrans(path)
