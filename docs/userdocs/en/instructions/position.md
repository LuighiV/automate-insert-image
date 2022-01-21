# How the image position is calculated

Apart from specifying the width and height of the image, the software has the option to set the *LeftOffset* and *BottomOffset* which are relative to the reference position which is set to the top corner of the text  or to the absolute position of the corner page.

The characteristics for each option of positioning are:

## Reference relative to text

When the reference is relative to text, and the entered text is found in the document, the reference point is set to the top left corner of the first character of the placeholder. Then, relative to this position, the image is positioned in the page. If there are various places where the text is found, the image is inserted only in the first occurrence.

![Reference relative to text](~/images/text_reference.png "Reference relative to text")

## Reference relative to page

In this case the reference is set to an absolute position in the page which is one of its four corners. Then, relative to this reference point, the software inserts the image.

![Reference relative to page](~/images/page_reference.png "Reference relative to page")