import d3dshot

d = d3dshot.create()
d.display = d.displays[3]
f = open('vF.txt', 'r')
vF = f.read()
f2 = open('vP.txt', 'r')
vP = f2.read()
vF2 = str(vF + '.png')
d.screenshot_to_disk(directory=vP, file_name=str(vF + '.png'))
f.close()
f2.close()