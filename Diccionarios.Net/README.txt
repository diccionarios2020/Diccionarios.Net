Diccionarios.Net
================

This is the source code for Diccionarios.  Not the original 2009 source code: but produced 
using a disassembler and some resource-reading tools.  The user interface was recreated 
manually but should be identical.

The original release code contained some code for an options form.  This was not accessible
in the release version, so I was unable to redraw the screens.  This material has been
omitted.  Likewise there was a debug form of some sort, also omitted.

The greyed-out buttons on the toolbar are as per the original.  It looks as if there was
some kind of multi-page select planned, but never completed.  I have not attempted to include
any of this.

This code was compiled in VB.Net 2010.  It could probably be back-ported to 2008, but I have
not tried.

To run the code, you need to have some dictionaries available.  The code looks for the folder
"Diccionarios" in the same folder as the .exe.  One trick that I imported into
the code base from a previous version is to create a folder, "D:\Proyecto Libros", and place
a directory "Diccionarios" inside that.  If there is no folder under the exe, it looks there.

In the "Diccionarios" folder, create a sub-folder for each dictionary, containing the page
images: Glare, Liddel-Scott, etc.  

Roger Pearse
30th March 2020