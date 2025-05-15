# Plant3dSimpleClipBox - Sample Code: use at own risk!
Clipping the drawing area for easier access to the relevant parts for the 3D designer

Clipbox is a well known tool for many 3D applications. Attached you see an approach to do the same thing in Plant 3D.
It is trying to solve the task in the most simple way. It is using the "isolate" Plant command and the "xclip" AutoCAD command.
That's why in shaded mode, if you xclip xrefs you will get problems with the snaps, so you need to go to wireframe to get the expected snaps.
You can also search for the DWgs that have parts with collisions with the clipbox. This works with the AutoCAD "boundaries". 
Be aware of what these boundaries are to understand what drawings will be attached. 
Code should work for 2016 and higher (tested only on 2016, 2017).
This tool is example code to show you how you could work more efficiently. It is not supported by Autodesk and you use it at own risk:

# Manual for ClipBox.dll for Plant3d

<b>command: preparebox</b>
action: box visible preparebox selection = clipbox at selection
            box invisible preparebox selection = make visible last used clipbox
            box visible preparebox noselection = end clipping, make clipbox invisible
Known issues: If you find yourself suddenly in the middle of a xlcip command, then you didn’t exactly follow the intended workflow. In this case just click “Esc” several times. Use orthomode off (F8 to switch on/off), better for resizing the box
Update: you can now select 2 objects to define the box

<b>command: clipping</b>
action: box prepared clipping = clip the box
Known issues: Depending on how the xclip command was used before, things might behave different (e.g. the outside of the box will be clipped instead of the inside). In this case execute the xclip command manually, the system will remember the settings.

<b>command: loadRequiredXrefs</b>
prerequisites: In the project, ALWAYS use overlay when attaching xrefs! Detach all xrefs.
action: box prepared loadRequiredXrefs = find required xrefs and load them.
Known issues: Possibly proxy objects from other programs (zombie objects) will disturb the program, has to be tested. Just “Plant 3d Drawings” are searched, not the “Related Files”
Update: there is some sort of caching now to speed up the command, so there will be a .bounds file per DWG, you will also be asked for a path, but you can just click enter on a Plant project

<h2>License</h2>
This sample is licensed under the terms of the <a href="http://opensource.org/licenses/MIT">MIT License</a>. Please see the <a href="https://github.com/Henaccount/Plant3dSimpleClipBox/blob/master/LICENSE">LICENSE</a> file for full details.

