# InCite

![logo_outline](https://user-images.githubusercontent.com/65059714/233111424-71272d58-3b6b-47b6-9c68-0dab383a0e7a.png)

Paid for citation tools work, but they can be inconsistent, require syncing, and of course cost money, with different institutions using different tools. Microsoft Word already has citations built into it, so I made a python script which could take downloaded citation files (nbib, enw, ris), convert them into a format that word could use, and inject them into the active Word document using pywin32 and COMs. I also coded a GUI for enabling easier search through the references that are already in the document. I tried to import as little as possible to make it as self contained, yet highly functional as possible. This is a fun side project and work in progress, so please fork it and suggest changes if you'd like to assist in any way. I'm pretty new to sharing my code so hopefully I'm doing everything right.

# Use
The script can be used directly with python. The program can be used with minimal terminal interfacing. In either case, it can open citation files to format and insert them into an active word document, or it can be opened standalone to show a GUI for browsing and inserting citations from the current list in the active Word document.

# Compiled version for windows
The first release includes a compiled exe for use on windows. You can set it to be the default program for opening citation formats from journal websites. You can also simply open it, with a word document open and it'll bring up the interface for inserting a citation already in the word documents current list. 

# Work in progress
- [ ] Make executable for Mac

- [ ] More citation file formats?

- [ ] More thoroughly validate how different journals use each format

- [ ] Convert other library formats to Word Citation format (an XML variant)

- [ ] Make an Add-on so it is self-contained (can't do this alone)

- [x] Get a cute snake logo
