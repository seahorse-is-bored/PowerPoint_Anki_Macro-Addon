# PowerPoint Anki Macro & Addon
A macro you can use for powerpoint to import straight into Anki. 

Some shit code I made (only works on MacOS)
(Please excuse the horrible code, I haven't coded in a while and I've never used VBA)
Currently you have to format the notes section like a CSV, but if anyone can be bothered to improve it you are more than welcome to do so.
There's 2 parts to this, a macro which you add to the powerpoint file and an addon which you add to the addons folder

To export all the notes form powerpoint, just press the run macro.

To import into anki go to tools -> Import from PowerPoint

**Macro**
You need to add this to your PowerPoint file when you save it locally, or you can make a template or something - all it does is export the notes and the file path to a location called ~/.anki-PPT

you may have to grant access to powerpoint

**Addon**
Create an addon folder (mine was called 123_321) and add the file __init__.py
