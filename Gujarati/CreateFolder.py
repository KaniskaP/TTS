import os
set = 'Set 1'
root_dir = os.getcwd() + '\Audio\\' + set
subfolders = ('\Menus', '\Paragraph', '\Sentences\MOS', '\Sentences\DMOS')
i = 0
for folder in subfolders:
    os.makedirs(root_dir + subfolders[i])
    i+=1
