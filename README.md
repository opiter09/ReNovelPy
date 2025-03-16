# ReNovelPy
This is a tool to turn Ren'Py Visual Novels into regular novels (i.e. Word documents). That way, you can print them out, and then
when climate change destroys civilization these VNs won't all be lost to time.

NOTE 1: This tool is is only designed for Windows (unless you want to try running the Python yourself, that is). For Mac and Linux, I
can only point you to WINE (https://www.winehq.org/).

NOTE 2: If you cannot or do not want to pay for Microsoft Word, these documents can be edited for free with Google Docs, or with
offline alternatives such as LibreOffice (https://www.libreoffice.org/).

In terms of the visuals, obviously most things are not saved: you will get a list of characters with one sprite each at the beginning,
and then backgrounds, CGs, effects, etc. intermittently. However, for some games even this can create files of very large sizes, and
obviously it would be rather impratical to print out a great many images. Therefore, there are options to include character sprites and
selected images [1], only character sprites, or no images at all. If you choose to do so, it will give you the file name as a "description"
for excluded images.

Similarly, sound effects have their file names shown, and the chapter headings are the internal "label" names. As a result of this, along
with many other oddities, please take advantage of this producing an editable Word doc to clean things up.

Furthermore, this was designed with kinetic (choiceless) VNs in mind. There is technically a system to handle choices (it displays them to
you and you pick one each time), but this does not take variables into account, so only the simplest frameworks will work. Making
variables work is really hard (to me at least), and since I figure that choice-heavy VNs are less suited to this conversion anyway, 
this is where we're at. However, the code is in the Public Domain, so if you would like to improve on it, by all means do so.

Finally, please make sure to always fairly compensate VN developers for their work. I'm not really sure why someone would pirate
an inherently inferior book-conversion of a VN instead of the VN itself, but I would still hate to find out that someone lost sales
because of this utility. Remember, VN developers are people too, and they need money to survive just as much as you do.

On a similar note, I recently played a VN with the restriction in its Readme that "ripping the game's graphical or musical assets is
not permitted." Please do not use ReNovelPy on a VN with this restriction, or any other restriction that would apply.

[1]: The code is very flexible when it comes to location: you can choose images at the site of the game's executable, or at the site of
novel.exe; inside of a translation folder, or out in the main one. However, it is important to be aware that for games with archives,
you will only be able to choose the files by going into the folder where novel.exe is, since that is where they have been unpacked.