Photolog_Desktop by Dan (dhaggerty@pronetgroup.com)

A very brief explanation of how the app works.

Photolog_Desktop can simply be seen as an assembly area for the files on your system. It does nothing to the underlying files. It does not move them or rotate them. 

When you drag images into the assembly area the list displays a thumbnail of the image and the last 4 digits of the file name. 
There is also a caption of text for you to fill in.
Clicking on an image updates the large picture box to show the image. Here you can rotate the image.

When you save a project, a simple temorary file is made (.xml). 
It is recommended you store the .xml file in same folder as your images. 
This is handy especially if you copy the folder of images to a thumb drive or other computer to work on as you can use the "Change parent folder option" to recover all your work.

This .xml file contains 3 key pieces of information per image:
1. The location of the image on your computer/drive. eg, "C:\Users\dhaggerty\Documents\images\bayou.jpg"
2. The caption you enetered for that image "This is the Bayou"
3. The rotation (stored as on of 4 integers) that you applied to that image. Either "0", "1", "2", "3".

When you resume a project, using the example above, Photolog_Desktop does the following:
1. Looks for the image at "C:\Users\dhaggerty\Documents\images\bayou.jpg"
2. Loads the caption
3. Rotates the image by the integer stored

At no point does Photolog_Desktop rotate or alter the images actually stored on your computer. 
