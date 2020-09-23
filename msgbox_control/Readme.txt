An Alternative Message Box (with designer)
Written by Kevin Figg

The following code from other people has been used.
1. The original alternative message box by 
2. chameleonButton Button by Conchuki

To use this messagebox you need the following files added to your project. 

1. Form 'frmmsgbox' (file: frmMSGBOX.frm and frmMSGBOX.frx)
2. Module 'message' (file: message.bas)
3. user Control 'chameleonButton' (file: chameleonButton.ctl and chameleonButton.ctx)

You might want to copy the files to your project folder first.

To get a feel of the message box I've added a designer which will a) help you get used to the options, and b) see the message box befoe you add it to your project. If you want you can compile the demo app, and use the designer each time you want to create a msgbox. But remember you need the above files in your project for it to work.

The options is the demo project speak for themselves.

By Saving as default, each time the app is loaded, the save options are loaded. This is
freat if you are working on a large project where you can set the msgbox options to a
uniform then just edit the message.


Some Issues
-----------
I have no idea how to make the msgbox vbsystemmodal

The text is selectable even though the mouse pointer doesn't change to an 'I'.
This also means the text can be copied and pasted which isn't a bad thing. The text
is of course read only, but if I stop the selection of text by making it disabled, the text
will turn gray.



Comments to Kevinfigg@hotmail.com please (and on PSC of course)

---
Thanks for downloading it

Kevin Figg