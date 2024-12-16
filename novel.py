import os
import shutil
import subprocess
import sys
import FreeSimpleGUI as psg
import docx # note: the module is named "python-docx" in pip
import unrpa

def sentenceCase(string):
    temp = string.replace("_", " ").replace("-", " ")
    temp2 = ""
    for s in temp.split(" "):
        if (len(s) > 1):
            temp2 = temp2 + s[0].upper() + s[1:]
        else:
            temp2 = temp2 + s[0].upper()
    return(temp2)

path = psg.popup_get_file("Game Executable:", font = "-size 12")
fo = path.split("/")[-2]
fi = path.split("/")[-1]
folder = "./game_" + fo.replace(" ", "_") + "/"
# print(folder)

if (os.path.exists(folder) == False):
    shutil.copytree(path[0:-(len(fi))] + "game", folder)
    for root, dirs, files in os.walk(folder):
        for file in files:
            if (file.endswith(".rpa") == True):
                rpa = unrpa.UnRPA(filename = os.path.join(root, file), path = folder)
                rpa.extract_files()
                os.remove(os.path.join(root, file))

labels = ["start"]
f = open(folder + "screens.rpy", "rb")
r = f.read().decode("UTF-8", errors = "ignore") # python is weird about Japanese etc. when using "rt"
f.close()
for l in r.split("\n"):
    if ('action Start("' in l):
        labels.append(l.split('action Start("')[1].split('"')[0])
# print(labels)

title = ""
f = open(folder + "options.rpy", "rb")
r = f.read().decode("UTF-8", errors = "ignore")
f.close()
for l in r.split("\n"):
    if ("config.name" in l):
        title = l.split('"')[1]
        break
    elif ("config.window_title" in l):
        title = l.split('"')[1]
        break
titleU = title.upper()
# print(title)

combined = bytes(0) # to avoid looping through files all the time
for root, dirs, files in os.walk(folder):
    for file in files:
        if (file.endswith(".rpy") == True):
            f = open(os.path.join(root, file), "rb")
            r = f.read()
            f.close()
            combined = combined + r + ("\n\n").encode("UTF-8", errors = "ignore")
combined = combined.decode("UTF-8", errors = "ignore")
# print(combined[0:50])
combL = list(combined.split("\n")).copy()
combL = [x.strip() for x in combL]

new = docx.Document()
head = new.add_heading(titleU, 0)
head.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
new.add_page_break()


new.add_heading("DRAMATIS PERSONÆ", 1)
nameVars = {}
usedNames = {}
for i in range(len(combL)):
    l = combL[i]
    if ('Character("' in l):
        name = l.split('Character("')[1].split('"')[0]
        for j in range(len(name)):
            if ((j < (len(name) - 1)) and (name[j] == "}")):
                if (name[j + 1] != "{"):
                    name = name[(j + 1):].split("{")[0]
                    break
        nameVars[l.split(" ")[1].split("=")[0]] = name
        image = ""
        if (l.endswith(")") == True):
            temp = l.replace(" =", "=").replace("= ", "=")
            if ('image="' in temp):
                image = temp.split('image="')[1].split('"')[0]
        else:
            j = i
            while (j < (len(combL) - 1)) and ('image="' not in combL[j].replace(" =", "=").replace("= ", "=")) and (('Character("' not in combL[j]) or (j == i)):
                j = j + 1
            temp = combL[j].replace(" =", "=").replace("= ", "=")
            if (('Character("' not in temp) and ('image="' in temp)):
                image = temp.split('image="')[1].split('"')[0]
        sprite = ""
        if (image != ""):
            # print(image)
            for k in range(len(combL)):
                l2 = combL[k]
                if ((l2.startswith("image ") == True) and ((" " + image + " ") in l2)):
                    if (('.png"' in l2) or ('.jpg"' in l2)):
                        m = k
                    else:
                        m = k
                        while (m < (len(combL) - 1)) and ('.png"' not in combL[m]) and ('.jpg"' not in combL[m]):
                            m = m + 1
                    for small in combL[m].split('"'):
                        if ((small.endswith(".png") == True) or (small.endswith(".jpg") == True)):
                            sprite = folder + small
                            # print(sprite)
                    break
        if (((name not in usedNames.keys()) or (usedNames[name] == "")) and (name.replace("?", "") != "")):
            if ((sprite not in usedNames.values()) or (sprite == "")):
                usedNames[name] = sprite
# print(usedNames)

for n in usedNames.keys():
    new.add_heading(n, 1)
    if (usedNames[n] != ""):
        new.add_picture(usedNames[n])
new.add_page_break()
# new.save("./" + title.replace(" ", "_") + ".docx")
            
curr = ""
for lab in labels:
    curr = lab
    while ("label " + curr + ":\n") in combined:
        new.add_header(sentenceCase(curr), 0)
        ind = combL.index("label " + curr + ":\n")


    

