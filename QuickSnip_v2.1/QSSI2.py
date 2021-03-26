import d3dshot
import os

def getScreen():
    vPx = os.path.split(vF2)
    vP3 = vPx[0] # split new vPathFull variable into new dir & new file_name
    vF3 = vPx[1]
    d = d3dshot.create()   # take screenshot & save using vP2 as dir & vF3 as filename
    d.display = d.displays[2]
    d.screenshot_to_disk(directory=vP3, file_name=vF3)
    f.close()
    f2.close()

# open vF.txt save contents as variable vF
f = open('vF.txt', 'r')
vF = str(f.read() + '.png')
# open vP.txt save contents as variable vP
f2 = open('vP.txt', 'r')
vP = f2.read()
# check if path exists, if it does increase iterator
vF2 = str(vP + vF)
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