# -*- coding: utf-8 -*-
"""
@author: liuzhiwei

@Date:  2020/11/15
"""

import os
import logging

import win32com.client

logger = logging.getLogger('Sun')
logging.basicConfig(level=20,
                    # format="[%(name)s][%(levelname)s][%(asctime)s] %(message)s",
                    format="[%(levelname)s][%(asctime)s] %(message)s",
                    datefmt='%Y-%m-%d %H:%M:%S'  # 注意月份和天数不要搞乱了，这里的格式化符与time模块相同
                    )


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


if __name__ == "__main__":
    transDict = {}
    transDict.update({1: {'name': 'pptx', 'formatType': 11}})
    transDict.update({2: {'name': 'ppt', 'formatType': 1}})
    transDict.update({3: {'name': 'pdf', 'formatType': 32}})

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
            op = pptTrans(infoDict, path)
