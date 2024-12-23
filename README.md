# ReNovelPy
This is a tool to turn Ren'Py Visual Novels into regular novels (i.e. Word documents). That way, you can print them out, and then
when climate change destroys civilization these VNs won't all be lost to time.

NOTE: This tool is is only designed for Windows (unless you want to try running the Python yourself). For Mac and Linux, I can only
point you to WINE: https://www.winehq.org

In terms of the visuals, obviously most things are not saved: you will get a list of characters with one sprite each at the beginning,
and then backgrounds, CGs, effects, etc. intermittently. However, for some games even this can creates files of very large sizes, so
there is an option to only include the character sprites, or no visuals at all. If so, it will give you the file name as a "description."

Similarly, sound effects have their file names shown, and the chapter headings are the internal "label" names. As a result of this, along
with many other oddities, please take advantage of this producing an editable Word doc to clean things up.

Furthermore, this was designed with kinetic (choiceless) VNs in mind. There is technically a system to handle choices (it displays them to
you and you pick one each time), but this does not take variables into account, so only the simplest frameworks will work. Making
variables work is really hard (to me at least), and since I figure that choice-heavy VNs are less suited to this conversion anyway, 
this is where we're at. However, the code is in the Public Domain, so if you would like to improve on it, by all means do so.

Finally, please, please, PLEASE do not use this to facilitate piracy. It does not come with any game's assets, but it does unpack
all archives present, and of course you could share the Word doc around if you wanted to. None of this would be possible without the
hard work and dedication of all the VN devs out there, and they deserve fair compensation for their efforts.