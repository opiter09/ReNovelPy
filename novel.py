import os
import re
import shutil
import subprocess
import sys
import FreeSimpleGUI as psg
import docx # note: the module is named "python-docx" in pip
import unrpa

def titleCase(string):
    temp = string.replace("_", " ")
    if (" " not in string):
        temp = temp.replace("-", " ")
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
combined = combined.replace("\\[", "[").replace("\\]", "]")
combined = combined.replace("\\{", "<").replace("\\}", ">")
# print(combined[0:50])
combL = list(combined.split("\n")).copy()
combL = [x.strip() for x in combL]

def findSprite(image):
    global combL
    global folder

    if (image == ""):
        return("")

    sprite = ""
    for i in range(len(combL)):
        l2 = combL[i]
        if ((l2.startswith("image ") == True) and ((" " + image + " ") in l2)):
            if (('.png"' in l2) or ('.jpg"' in l2)):
                j = i
            else:
                j = i
                while ((j < (len(combL) - 1)) and ('.png"' not in combL[j]) and ('.jpg"' not in combL[j])) or (combL[j][0] == "#"):
                    j = j + 1
            for small in combL[j].split('"'):
                if ((small.endswith(".png") == True) or (small.endswith(".jpg") == True)):
                    sprite = folder + small
                    # print(sprite)
            break
    return(sprite)
    
new = docx.Document()
head = new.add_heading(titleU, 0)
head.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
new.add_page_break()


new.add_heading("DRAMATIS PERSONÆ", 1)
nameVars = {}
usedNames = {}
for i in range(len(combL)):
    l = combL[i]
    if (('Character("' in l) and (l[0] != "#")):
        name = l.split('Character("')[1].split('"')[0]
        for j in range(len(name)):
            if ((j < (len(name) - 1)) and (name[j] == "}")):
                if (name[j + 1] != "{"):
                    name = name[(j + 1):].split("{")[0]
                    break
        var = l.split(" ")[1].split("=")[0]
        if (var[-1] == " "):
            var = var[0:-1]
        nameVars[var] = name
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
            sprite = findSprite(image)
        if (((name not in usedNames.keys()) or (usedNames[name] == "")) and (name.replace("?", "") != "")):
            if ((sprite not in usedNames.values()) or (sprite == "")):
                usedNames[name] = sprite
# print(usedNames)

for n in usedNames.keys():
    new.add_heading(n, 1)
    if (usedNames[n] != ""):
        new.add_picture(usedNames[n])
new.add_page_break()

otherImageVars = []
for l in combL:
    if (l.startswith("image ") == True):
        temp = l[6:].split(":")[0].split("=")[0]
        if (temp[-1] == " "):
            temp = temp[0:-1]
        if (temp not in nameVars.keys()):
            otherImageVars.append(temp)
otherImageVars.sort() # shorter names come first

p = ""
def handleTags(string):
    global p
    
    tags =  re.split(r'([\{\}])', string) # thank you Stack Exchange
    runs = []
    for t in tags:
        runs.append([t])
    for i in range(len(tags) - 2):
        if (tags[i] == "{"):
            runs[i].append("g" * 1000)
            runs[i + 1].append("g" * 1000)
            runs[i + 2].append("g" * 1000)
            if (tags[i + 1][0] != "/"):
                safe = tags[i + 1].replace(" ", "")
                # b is bold, i is italics, s is strikethrough, u is underline, plain removes all 4, space is x spaces, size does math
                if (safe.split("=")[0] in ["b", "i", "s", "u", "plain", "size"]):
                    for j in range(i + 1, len(tags)):
                        runs[j].append(safe)
                elif (safe.split("=")[0] == "space"):
                    runs[i + 1][0] = " " * int(safe.split("=")[1])
            else:
                safe = tags[i + 1].replace(" ", "")
                for j in range(i + 1, len(tags)):
                    runs[j].remove(safe[1:])
    for r in runs:
        if ((r[0] != "") and ((("g" * 1000) not in r) or (r[0].replace(" ", "") == ""))):
            thing = p.add_run(r[0])
            thing.font.size = docx.shared.Pt(11)
            if ("b" in r):
                thing.bold = True
            if ("i" in r):
                thing.italic = True
            if ("s" in r):
                thing.strike = True
            if ("u" in r):
                thing.underline = True
            if ("plain" in r):
                thing.bold = False
                thing.italic = False
                thing.strike = False
                thing.underline = False
            for val in r:
                if (val.startswith("size") == True):
                    if ("*" in val):
                        thing.font.size = int(thing.font.size * float(val.split("*")[1]))
                    elif ("+" in val):
                        thing.font.size = thing.font.size + docx.shared.Pt(int(val.split("+")[1]))
                    elif ("-" in val):
                        thing.font.size = thing.font.size - docx.shared.Pt(int(val.split("-")[1]))
                    else:
                        thing.font.size = docx.shared.Pt(int(val.split("=")[1]))      

curr = ""
for lab in labels:
    curr = lab
    while ("\n" + "label " + curr + ":") in combined:
        new.add_heading(titleCase(curr), 0)
        ind = 0
        for i in range(len(combL)):
            if (combL[i].startswith("label " + curr + ":") == True):
                ind = i
                break
        while True:
            line = combL[ind]
            if (line == ""):
                ind = ind + 1
                if (ind == len(combL)):
                    curr = "g" * 1000
                    break
                else:
                    continue
            if ((line.startswith("play sound ") == True) or (line.startswith("play audio ") == True)):
                p = new.add_paragraph()
                p.add_run("Sound: " + titleCase(line.split('"')[1].split("/")[-1][0:-4])).italic = True
            elif ((line.startswith("scene bg ") == True) or (line.startswith("scene cg ") == True)):
                temp = ""
                for v in otherImageVars:
                    if ((" " + v + " ") in line[9:]):
                        temp = v # don't break so longer names trump shorter ones
                sprite = findSprite(temp)
                if (sprite != ""):
                    new.add_picture(sprite)
            elif (line.startswith("scene ") == True):
                temp = ""
                for v in otherImageVars:
                    if ((" " + v + " ") in line[6:]):
                        temp = v # don't break so longer names trump shorter ones
                sprite = findSprite(temp)
                if (sprite != ""):
                    new.add_picture(sprite)
            elif (line.startswith("show ") == True):
                temp = ""
                for v in otherImageVars:
                    if ((" " + v + " ") in line[5:]):
                        temp = v # don't break so longer names trump shorter ones
                sprite = findSprite(temp)
                if (sprite != ""):
                    new.add_picture(sprite)
            elif (line.startswith("jump ") == True):
                curr = line[5:]
                break
            elif (line.startswith("show text ") == True):
                 p = new.add_paragraph()
                 handleTags(line.split('"')[1])
            elif (line[0:6] == "return"):
                curr = "g" * 1000
                break
            elif (line[0] == '"'):
                p = new.add_paragraph()
                handleTags(line.split('"')[1])
            else:
                theKeys = list(nameVars.keys()).copy()
                theKeys.sort() # shorter names come first
                temp = ""
                for k in theKeys:
                    if ((line.startswith(var + " ") == True) or (line.startswith(var + '"') == True)):
                        temp = k # don't break so longer names trump shorter ones
                if (temp != ""):
                    p = new.add_paragraph()
                    if (nameVars[temp] != ""):
                        p.add_run(nameVars[temp] + ": ").bold = True
                    handleTags(line.split('"')[1])
            ind = ind + 1
            if (ind == len(combL)):
                curr = "g" * 1000
                break

new.save("./" + title.replace(" ", "_") + ".docx")