# Plant-3D-Hoses

See also video: https://youtu.be/2-rJcDCbhmw


How to create hoses in Plant 3d? You can create them using a spline and extruding a circle on it. Then you make a block out of it. Then you prepare it using PLANTPARTCONVERT. Then you use it as a custom part using "add custom part" from the dynamic tool palette. Then you write the properties in the properties palette..

Exhausted? Try the API approach to automate it. Sample code is attached. It will create the part based on the spline, read properties from the spec (spec cannot be open in the spec editor at the same time) and calculate the length write it in the properties (length is not accurate I think)

I created this as a test and it worked on Plant 3D 2016 (usually this works on the newer versions then as well), shape 3 is the most reliable one (other do not always work) and you need to go back from the vertices mode in the spline to the fit mode before applying the command. Alternative spec path might not work.

See more details from the following manual and the video in the attachment.

 

Manual for hose.dll:

 

##how to use:

-----------

1. Draw Spline (see mp4 manual how to draw spline for hose)

2. call:  command hose "shortdesc=Myhose,shape=1,scale=1,specalternatepath=C:/not/default/place/for/spec sheets"  

   (put it exactly like this in user defined command, then click it from palette or ribbon)

3. place hose and connect it 

 

 

##how to set up:

--------------

-produce your hoses in the catalog then add them to the spec sheets, type is "coupling", geomety doesn't matter, 

 because it is not used, just the information is read from the spec

-configure skey for the hose in the IsoSkeyAcadBlockMap.xml, create block if not present

-short description and nominal diameter is the selection criteria, so if you have different hoses, 

 use different short descriptions. Add this short description to the command (see above)

 

 

##limitations:

------------

-works on Plant 3D 2016 (most likely the upper versions as well)

-encodes are the same on both ends

 

 

##parameters:

-----------

shortdesc: see above

shape: 4 different shapes 0,1,2,3, 3 is most simple

scale: you can scale the diameter

specalternatepath: (can be empty if not needed) if the specs are not in the default project folder, you can specify here (you have to use / for the path, no \ allowed)

 

See this article about how do create the dll and how to install it:  http://autode.sk/2jYKHJy
