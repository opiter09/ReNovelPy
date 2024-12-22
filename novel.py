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

combined = "" # to avoid looping through files all the time
for root, dirs, files in os.walk(folder):
    for file in files:
        if (file.endswith(".rpy") == True):
            f = open(os.path.join(root, file), "rb")
            r = f.read()
            f.close()
            combined = combined + "\n\n" + r.decode("UTF-8", errors = "ignore")
combined = combined.replace("\\[", "[").replace("\\]", "]")
combined = combined.replace("\\{", "<").replace("\\}", ">")
weirdData = [
    (b"\xef\xbb\xbf").decode("UTF-8", errors = "ignore")
]
for x in weirdData:
    combined = combined.replace(x, "")
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
        l2 = combL[i].replace(" =", "=").replace("= ", "=")
        if ((l2.startswith("image ") == True) and (((" " + image + " ") in l2) or ((" " + image + "=") in l2))):
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

section = new.sections[0]
pageWidth = section.page_width - section.left_margin - section.right_margin


new.add_heading("Dramatis Personæ", 0)
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
        new.add_picture(usedNames[n], width = (pageWidth * 0.5))
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
# print(nameVars.keys())
# print(otherImageVars)

p = ""
def handleTags(string):
    global p
    
    tags =  re.split(r'([\{\}])', string) # thank you Stack Exchange
    runs = []
    for t in tags:
        runs.append([t])
    lastSize = ""
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
                    if (safe.split("=")[0] == "size"):
                        lastSize = safe
                elif (safe.split("=")[0] == "space"):
                    runs[i + 1][0] = " " * int(safe.split("=")[1])
            else:
                safe = tags[i + 1].replace(" ", "")
                for j in range(i + 1, len(tags)):
                    try:
                        runs[j].remove(safe[1:])
                    except ValueError:
                        if (safe == "/size"):
                            runs[j].remove(lastSize)
                            
    for r in runs:
        if ((r[0] != "") and ((("g" * 1000) not in r) or (r[0].replace(" ", "") == ""))):
            thing = p.add_run(r[0])
            thing.font.size = docx.shared.Pt(12)
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
        # print(curr)
        new.add_heading(titleCase(curr), 0)
        ind = 0
        for i in range(len(combL)):
            if (combL[i].startswith("label " + curr + ":") == True):
                ind = i
                break
        while True:
            line = combL[ind]
            if ((line == "") or (line.startswith("#") == True) or (line.startswith("$") == True)):
                ind = ind + 1
                if (ind == len(combL)):
                    curr = "g" * 1000
                    # print("finalBlank")
                    break
                else:
                    continue
            if ((line.startswith("play sound ") == True) or (line.startswith("play audio ") == True)):
                p = new.add_paragraph()
                p.add_run("Sound: " + titleCase(line.split('"')[1].split("/")[-1][0:-4])).italic = True
            elif ((line.startswith("scene bg ") == True) or (line.startswith("scene cg ") == True)):
                temp = ""
                for v in otherImageVars:
                    if ((" " + v) in line[8:]):
                        temp = v # don't break so longer names trump shorter ones
                sprite = findSprite(temp)
                if (sprite != ""):
                    new.add_picture(sprite, width = pageWidth)
            elif (line.startswith("scene ") == True):
                temp = ""
                for v in otherImageVars:
                    if ((" " + v) in line[5:]):
                        temp = v # don't break so longer names trump shorter ones
                sprite = findSprite(temp)
                if (sprite != ""):
                    new.add_picture(sprite, width = pageWidth)
            elif False: # elif ((line.startswith("show ") == True) and (line.startswith("show text ") == False)):
                temp = ""
                for v in otherImageVars:
                    if ((" " + v) in line[4:]):
                        temp = v # don't break so longer names trump shorter ones
                sprite = findSprite(temp)
                if (sprite != ""):
                    new.add_picture(sprite, width = pageWidth)
            elif (line.startswith("jump ") == True):
                curr = line[5:]
                # print("jump")
                break
            elif (line.startswith("show text ") == True):
                 p = new.add_paragraph()
                 handleTags(line.split('"')[1])
            elif (line[0:6] == "return"):
                curr = "g" * 1000
                # print("finalReturn")
                break
            elif (line[0] == '"'):
                p = new.add_paragraph()
                handleTags(line.split('"')[1])
            else:
                theKeys = list(nameVars.keys()).copy()
                theKeys.sort() # shorter names come first
                temp = ""
                for k in theKeys:
                    if ((line.startswith(k + " ") == True) or (line.startswith(k + '"') == True)):
                        temp = k # don't break so longer names trump shorter ones
                if (temp != ""):
                    p = new.add_paragraph()
                    if (nameVars[temp] != ""):
                        name = p.add_run(nameVars[temp] + ": ")
                        name.bold = True
                        name.font.size = docx.shared.Pt(12)
                    handleTags(line.split('"')[1])
            ind = ind + 1
            if ((ind == len(combL)) and (line.startswith("jump ") == False)):
                curr = "g" * 1000
                # print("finalNormal")
                break
    # print(curr + " out")

new.save("./" + title.replace(" ", "_") + ".docx")