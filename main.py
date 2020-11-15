# -*- coding: utf-8 -*-
"""
@author: liuzhiwei

@Date:  2020/11/15
"""

import os
import logging

from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
import win32com.client

logger = logging.getLogger('Sun')
logging.basicConfig(level=20,
                    # format="[%(name)s][%(levelname)s][%(asctime)s] %(message)s",
                    format="[%(levelname)s][%(asctime)s] %(message)s",
                    datefmt='%Y-%m-%d %H:%M:%S'  # 注意月份和天数不要搞乱了，这里的格式化符与time模块相同
                    )


def getFiles(dir, suffix, ifsubDir=True):  # 查找根目录，文件后缀
    res = []
    for root, directory, files in os.walk(dir):  # =>当前根,根下目录,目录下的文件
        for filename in files:
            name, suf = os.path.splitext(filename)  # =>文件名,文件后缀
            if suf.upper() == suffix.upper():
                res.append(os.path.join(root, filename))  # =>吧一串字符串组合成路径
        if False is ifsubDir:
            break
    return res


class pptTrans:
    def __init__(self, infoDict, filePath):
        self.infoDict = infoDict
        self.filePath = filePath
        self.powerpoint = None

        self.init_powerpoint()
        self.convert_files_in_folder(self.filePath)
        self.quit()
        os.system('pause')

    def quit(self):
        if None is not self.powerpoint:
            self.powerpoint.Quit()

    def init_powerpoint(self):
        try:
            self.powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
            self.powerpoint.Visible = 2
        except Exception as e:
            logger.error(str(e))

    def ppt_trans(self, inputFileName):
        # https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype

        infoDict = self.infoDict
        formatType = infoDict['formatType']
        outputFileName = self.getNewFileName(infoDict['name'], inputFileName)

        if '' == outputFileName:
            return
        inputFileName = inputFileName.replace('/', '\\')
        outputFileName = outputFileName.replace('/', '\\')
        if '' == outputFileName:
            return
        if None is self.powerpoint:
            return
        powerpoint = self.powerpoint
        logger.info('开始转换：[{0}]'.format(inputFileName))
        deck = powerpoint.Presentations.Open(inputFileName)

        try:
            deck.SaveAs(outputFileName, formatType)  # formatType = 32 for ppt to pdf
            logger.info('转换完成：[{0}]'.format(outputFileName))
        except Exception as e:
            logger.error(str(e))
        deck.Close()

    def convert_files_in_folder(self, filePath):
        if True is os.path.isdir(filePath):
            dirPath = filePath
            files = os.listdir(dirPath)
            pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
        elif True is os.path.isfile(filePath):
            pptfiles = [filePath]
        else:
            self.logError('不是文件夹，也不是文件')
            return

        for pptfile in pptfiles:
            fullpath = os.path.join(filePath, pptfile)
            self.ppt_trans(fullpath)

    def getNewFileName(self, newType, filePath):
        try:
            dirPath = os.path.dirname(filePath)
            baseName = os.path.basename(filePath)
            fileName = baseName.rsplit('.', 1)[0]
            suffix = baseName.rsplit('.', 1)[1]
            if newType == suffix:
                logger.warning('文档[{filePath}]类型和需要转换的类型[{newType}]相同'.format(filePath=filePath, newType=newType))
                return ''
            newFileName = '{dir}/{fileName}.{suffix}'.format(dir=dirPath, fileName=fileName, suffix=newType)
            if os.path.exists(newFileName):
                newFileName = '{dir}/{fileName}_new.{suffix}'.format(dir=dirPath, fileName=fileName, suffix=newType)
            return newFileName
        except Exception as e:
            logger.error(str(e))
            return ''


class pngstoPdf:
    def __init__(self, infoDict, filePath):
        self.infoDict = infoDict
        self.powerpoint = None

        self.init_powerpoint()
        self.convert_files_in_folder(filePath)
        self.quit()
        os.system('pause')

    def quit(self):
        if None is not self.powerpoint:
            self.powerpoint.Quit()

    def init_powerpoint(self):
        try:
            self.powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
            self.powerpoint.Visible = 2
        except Exception as e:
            logger.error(str(e))

    def ppt_trans(self, inputFileName):
        # https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
        infoDict = self.infoDict
        formatType = infoDict['formatType']
        outputFileName = self.getNewFolderName(inputFileName)

        if '' == outputFileName:
            return ''
        inputFileName = inputFileName.replace('/', '\\')
        outputFileName = outputFileName.replace('/', '\\')
        if '' == outputFileName:
            return ''
        if None is self.powerpoint:
            return ''
        powerpoint = self.powerpoint
        logger.info('开始转换：[{0}]'.format(inputFileName))
        deck = powerpoint.Presentations.Open(inputFileName)

        try:
            deck.SaveAs(outputFileName, formatType)
            logger.info('转换完成：[{0}]'.format(outputFileName))
        except Exception as e:
            logger.error(str(e))
            return ''
        deck.Close()
        return outputFileName

    def convert_files_in_folder(self, filePath):
        if True is os.path.isdir(filePath):
            dirPath = filePath
            files = os.listdir(dirPath)
            pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
        elif True is os.path.isfile(filePath):
            pptfiles = [filePath]
        else:
            self.logError('不是文件夹，也不是文件')
            return

        for pptfile in pptfiles:
            fullpath = os.path.join(filePath, pptfile)
            folderName = self.ppt_trans(fullpath)
            try:
                self.png_to_pdf(folderName)
            except Exception as e:
                logger.error(str(e))
            for file in os.listdir(folderName):
                os.remove('{0}\\{1}'.format(folderName, file))
            os.rmdir(folderName)

    def png_to_pdf(self, folderName):
        picFiles = getFiles(folderName, '.png')
        pdfName = self.getFileName(folderName)

        '''多个图片合成一个pdf文件'''
        (w, h) = landscape(A4)  #
        cv = canvas.Canvas(pdfName, pagesize=landscape(A4))
        for imagePath in picFiles:
            cv.drawImage(imagePath, 0, 0, w, h)
            cv.showPage()
        cv.save()

    def getFileName(self, folderName):
        dirName = os.path.dirname(folderName)
        folder = os.path.basename(folderName)
        return '{0}\\{1}.pdf'.format(dirName, folder)

    def getNewFolderName(self, filePath):
        index = 0
        try:
            dirPath = os.path.dirname(filePath)
            baseName = os.path.basename(filePath)
            fileName = baseName.rsplit('.', 1)[0]

            newFileName = '{dir}/{fileName}'.format(dir=dirPath, fileName=fileName)
            while True:
                if os.path.exists(newFileName):
                    newFileName = '{dir}/{fileName}({index})'.format(dir=dirPath, fileName=fileName, index=index)
                    index = index + 1
                else:
                    break
            return newFileName
        except Exception as e:
            logger.error(str(e))
            return ''


if __name__ == "__main__":
    transDict = {}
    transDict.update({1: {'name': 'pptx', 'formatType': 11}})
    transDict.update({2: {'name': 'ppt', 'formatType': 1}})
    transDict.update({3: {'name': 'pdf', 'formatType': 32}})
    transDict.update({4: {'name': 'png', 'formatType': 18}})
    transDict.update({5: {'name': 'pdf(不可编辑)', 'formatType': 18}})

    hintStr = ''
    for key in transDict:
        hintStr = '{src}{key}:->{type}\n'.format(src=hintStr, key=key, type=transDict[key]['name'])

    while True:
        print(hintStr)
        transFerType = int(input("转换类型:"))
        if None is transDict.get(transFerType):
            logger.error('未知类型')
        else:
            infoDict = transDict[transFerType]
            path = input("文件路径:")
            if 5 == transFerType:
                pngstoPdf(infoDict, path)
            else:
                op = pptTrans(infoDict, path)
