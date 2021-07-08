import os
import d3dshot
import time

def takeScreenshot(vPath, vFilename):
    d = d3dshot.create()
    picPath = str(vPath + '/' + vFilename + '.png')
    if os.path.isfile(picPath):
        expand = 1
        while True:
            expand += 1
            vFilenameNew = vFilename.split('.png')[0] + str(expand)
            newPicPath = str(vPath + '/' + vFilenameNew + '.png')
            if os.path.isfile(newPicPath):
                continue
            else:
                vFilename = vFilenameNew
                break
    
    d.screenshot_to_disk(directory=vPath, file_name=str(vFilename + '.png'))