# InCite

![logo_outline](https://user-images.githubusercontent.com/65059714/232627309-c8725b29-1c49-42f6-9fa2-1172e22d4548.png)

Paid for citation tools work, but they can be inconsistent, require syncing, and of course cost money, with different institutions using different tools. Microsoft Word already has citations built into it, so I made a python script which could take downloaded citation files (nbib, enw, ris), convert them into a format that word could use, and inject them into the active Word document using pywin32 and COMs. I also coded a GUI for enabling easier search through the references that are already in the document. I tried to import as little as possible to make it as self contained, yet highly functional as possible. This is a fun side project and work in progress, so please fork it and suggest changes if you'd like to assist in any way. I'm pretty new to sharing my code so hopefully I'm doing everything right.

# Use
The script can be used directly with python, or it can be converted to an exe or executable by using pyinstaller (see below). Then the program can be used with minimual terminal interfacing. In either case, it can open citation files to format and insert them into an active word document, or it can be opened standalone to show a GUI for browsing and inserting citations from the current list in the active Word document.

# Compiled version for windows
Here is an exe that should work on Windows machines.
https://drive.google.com/file/d/1QpZaygRgkASX-0b7D7W2E3PR1x_eA7UF/view?usp=share_link
It can be set as the default program to open citation files, and when it does it will inject the citation into the active word document wherever the cursor is located (merging references where appropriate). If it is opened without a citation file, it will bring up the GUI which can be used to find already existing citations in the document and insert them wherever the cursor is.

# Work in progress
- [ ] Make executable for Mac

- [ ] More citation file formats?

- [ ] More thoroughly validate how different journals use each format

- [ ] Convert other library formats to Word Citation format (an XML variant)

- [ ] Make an Add-on so it is self-contained (can't do this alone)

- [x] Get a cute snake logo

- [ ] Decrease size of current exe (remove unnecessary dependencies or files)
