import os
from PIL import ImageGrab

f = open('vF.txt', 'r')
vF = f.read()
vF2 = str(vF + '.png')

def getScreen():
    img = ImageGrab.grab()
    img.save(vF2,'PNG')
    f.close()

if os.path.isfile(vF2):
    expand = 1
    while True:
        expand += 1
        new_vF2 = vF2.split(".png")[0] + str(expand) + ".png"
        if os.path.isfile(new_vF2):
            continue
        else:
            vF2 = new_vF2
            break

getScreen()