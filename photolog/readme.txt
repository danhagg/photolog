Photolog_Desktop v1.1
dhaggerty@pronetgroup.com

New in v1.1

Image Orientation. In the photolog app the thumbnail image, the Enlarged image and the image exported to Word should all be in the same orientation as the file saved in your computer system.

Vertical images in Word should now be the same height as horizontal images

The temp files bear the extension “.photolog” for clarity that they are associated with the photolog app. 
To protect users who may switch versions mid-project they can still open their .XML files with photolog. 

There is now a Warning on trying to exit the app that u are exiting and may lose unsaved work

Up/down keys changes which row you are viewing (and updates the picture box). It does NOT move the row.

Using mouse to click UP/DOWN buttons on page DOES move the ROW.

The Delete BUTTON on the app and on the keyboard deletes the selection.

You can know write text into a caption, hit ENTER, and immediately start writing into the following text box.

You can append multiple .photolog files together. For example, with three temp files, open the first temp file using “Resume”, append the second, then the third with “Append another temp file”. Make sure you append them in the order you want them published. 
Copying and pasting captions

It appears that using CTRL-C and CTRL-V to copy and paste captions around was causing tabs/whitespace to appear in the captions. This problem should be solved.

Insertion of blank captions into Word. This was caused by hidden "carriage return" characters at end of captions. Should be solved.

Small amount of horizontal whitespace added between thumbnails to separate images


INSTRUCTIONS v1.0
Photolog_Desktop can simply be seen as an assembly list for the image files on your system. 

The App does nothing to these underlying image files. The App does not move them or rotate them.

1. Open the app.
2. Open a folder of images.
3. Drag images from the folder into the LEFT panel (assembly list) of the app. These images actually STAY in the folder. A thumbnail is generated in the assembly list. 
4. You can drag multiple images into the assembly list at once.
5. Images already in the assembly list will be rejected if you try and drag them in but those not in the assembly list will be accepted.
6. You can left-click the mouse and drag over many items in assembly list to bulk select images
7. Images (individual or bulk-selection) can be moved using UP/DOWN button and deleted with DELETE button.
8. You right-click on the list and send all selected images to TOP or BOTTOM of the assembly.
9. Clicking on an image updates the large picture box to show the image.
10. There is a caption box you can fill in.
11. Click on a thumbnail to see an enlargement in the display box. Photos > 5MB will not be displayed.
12. The chart in bottom right displays all image sizes. Images > 2MB will have red bars and display their number in the assembly list above the bar.
13. The bar chart updates with the moving up/down and deleting of images. It also keeps a running total of the image sizes.
14. When you save a project, a simple temp file is made (.xml). It is recommended you store the .xml file in same folder as your images. 
15. "Save As". Use this the first time you save your project. The file name will then appear in the "Temp file" text box.
16. "Save". Will default save to the last place you saved. ie, the file indicated in the "Temp file" text box.
17. "Resume". Select the temp file to resume your project.
18. "Change Parent folder". AS LONG AS your temp file is stored in the same folder as the images... 
By selecting change parent folder and selecting your temp file, you should still be able to recover your work.

The 'temp' .xml file contains 2 key pieces of information per image:
A. The location of the image on your computer/drive. eg, "C:\Users\dhaggerty\Documents\images\bayou.jpg"
B. The caption you enetered for that image "This is the Bayou"

When you resume a project, using the example above, Photolog_Desktop does the following:
1. Looks for the image at "C:\Users\dhaggerty\Documents\images\bayou.jpg"
2. Loads the caption associated with that image.

If the Bayou image is no longer at "C:\Users\dhaggerty\Documents\images\bayou.jpg" because you are working from a different drive the you will get an error.
The "Change parent folder" option exists for this reason. Ensure the temp file is in the same folder as all your images.
Choose the "Change parent folder option".
Select temp file. And your work should repopulate.

19. PUBLISH. Creates a Word Document from your assembly list. You may find that, upon saving the Word document, the images are compressed. 
As the size of the Word Document should be significantly LOWER than the sum total of the images indicated in the bar chart.

