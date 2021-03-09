''' Template Auto Zooming Tool '''

import re  #  If you don't have the 're' library installed, please use pip to install it
import os
import sys
import datetime

class Temp:
    def __init__(self, x, y):
        self.x = x
        self.y = y

''' Configuration '''

new = Temp(200, 200) # Target Resolution
old = Temp(152, 152) # Source file resolution

zoom_font = 2

# Be sure to note the spaces in the middle or at the end of the following configurations
startx = '"start_x": '
starty = '"start_y": '
endx = '"end_x": '
endy = '"end_y": '
fonta = '"font_type": "$police '
fontan = '"font_type": "$police Narrow '
fontab = '"font_type": "$police Bold '
#fontai = '"font_type": "$police Italic '
fontsize = '"font_size": '

''' end '''

today = datetime.date.today()
file_data = ""
file = sys.argv[1]
f = open(file, 'r', encoding="utf_8")
line = f.readline()
if old.x > new.x and old.y > new.y:
    x = old.x / new.x
    y = old.y / new.y
else:
    x = new.x / old.x
    y = new.y / old.y
print(x,y)
while line:
    if line[0] != '#':
        nb = re.sub("\D", "", line)
        if old.x > new.x and old.y > new.y:
            if startx in line:
                line = startx + str(int(int(nb) / x)) + ',\n'
            if starty in line:
                line = starty + str(int(int(nb) / y)) + ',\n'
            if endx in line:
                line = endx + str(int(int(nb) / x)) + ',\n'
            if endy in line:
                line = endy + str(int(int(nb) / y)) + ',\n'
        else:
            if startx in line:
                line = startx + str(int(int(nb) * x)) + ',\n'
            if starty in line:
                line = starty + str(int(int(nb) * y)) + ',\n'
            if endx in line:
                line = endx + str(int(int(nb) * x)) + ',\n'
            if endy in line:
                line = endy + str(int(int(nb) * y)) + ',\n'
        if fontsize in line:
            line = fontsize + str(int(int(nb) + zoom_font)) + ',\n'
        
        # In this if statement below, make sure 'elif fonta in line:' is at the end
        if fontan in line:
            try:
                line = fontan + str(int(int(nb) + zoom_font)) + '",\n'
            except:
                line = line
        # elif fontai in line:
        #     try:
        #         line = fontai + str(int(int(nb) + zoom_font)) + '",\n'
        #     except:
        #         line = line
        elif fontab in line:
            try:
                line = fontab + str(int(int(nb) + zoom_font)) + '",\n'
            except:
                line = line
        elif fonta in line:
            try:
                line = fonta + str(int(int(nb) + zoom_font)) + '",\n'
            except:
                line = line
    file_data += line
    line = f.readline()
f.close()
filename = os.path.splitext(sys.argv[1])[0] + '_zoomed_' + today.strftime('%m%d') + os.path.splitext(file)[-1]
# If filename does not exist it will be created automatically, 
# 'w' means write data, before writing it will clear the original data in the file!
with open(filename,'w', encoding="utf_8") as f: 
    f.write(file_data)
print(file_data)
f.close()
