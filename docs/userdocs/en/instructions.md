# Instructions

This software was developed to insert an image in several MS Word / PDF files. In general, the process is similar for both cases, having some differences when chosing the document type and some options only available for a specific type. 

## Inserting image in MS Word files

For both, Word and PDF files there are two options to use the software: with individual files or with a folder which contain the files to be processed.

### For individual files:

1. Select *Word Document* as document input type.

2. Select *Files* as *Type*.

3. Click on *Explore* in *Files/folder Path* row to select a valid MS Word file or list of files. 

    ![Options to open Word file](~/images/word_file_1_2_3.png "Options to open Word file")

     It opens the file explorer  where you should localize the document and open it (you can also select various documents and insert the image in all of them).

    ![Open Word file(s)](~/images/word_file_3.png "Open Word file(s)")

4. Click on *Explore* in *Picture Path* row to select a valid image file to insert in the documents selected before. 

    ![Explore image file](~/images/word_file_4.png "Explore image path")

    It also opens the file explorer to select the desired picture:

    ![Open image file](~/images/word_file_4_1.png "Open image path")

    When the file is selected via the file explorer, an image preview is shown in the Preview Box. It also loads the dimensions measured in the selected unit, shown in the *Units* row.

5. You can modify the size of the image to insert and if *Keep aspect ratio* is checked, the other dimension is modified accordingly.
    You can also modify the offset from left (*LeftOffset*) and bottom (*BottomOffset*). You can also enter negative numbers to get an offset to the right and to the bottom. 

    ![Set dimensions](~/images/word_file_5.png "Set dimensions")

6. Once all the dimensions are set, you can enter the placeholder, which is a portion of text in the document which works as a reference.

    ![Set placeholder](~/images/word_file_6.png "Set placeholder")

    Keep in mind that the *LeftOffset* and *BottomOffset* are relative to the position of the top left corner of the text in the document.

7. In save options, you could select if you one to save the PDF in the same folder where it is the MS Word document or a subfolder and enter the name. If the folder doesn't exist, the software will create it for you.

    ![Choose location for saved PDF](~/images/word_file_7.png "Choose location for saved PDF")

8. Finally, you can click on *Generate* to process the documents and insert the image with the entered dimensions in the reference entered.

    ![All steps for Word files](~/images/word_file_8.png "All steps for Word files")

    When the process starts it opens a new dialog which shows in a progress bar the advance of the processing and prints informative messages about the current state of the processing flow.
    
    <img src="~/images/word_file_progress.png" alt="Progreso de procesamiento" width="70%"/>

Then, you can check the result opening the PDF. Don't forget to close the PDF, if you want to regenerate it, otherwise an error will occur as the software cannot write the file if it is open.

![Input vs output document files](~/images/word_file_diff.png "Input vs output document files")

### For a folder containing Word files:

It is pretty similar to the previous step changing in the second one:

1. Select *Word Document* as document input type.

2. Select *Folder* as *Type*.

3. Click on *Explore* in *Files/folder Path* row to select the folder containing the Word files.

    ![Options to open the folder](~/images/word_folder_1_2_3.png "Options to open the folder")

     It opens the folder explorer  where you should localize the folder that contains the Word files.

    <img src="~/images/word_folder_3.png" alt="Open folder" width="50%"/>

From here, it continues with the 4 to the 8 steps of individual files. When clicking in the *Generate* button it opens the progress dialog box which shows the status of processing for all of the Word documents inside the selected folder.

<img src="~/images/word_folder_progress.png" alt="Progreso de procesamiento" width="70%"/>

When opening the folder you can check the generated files (each one per Word document found in selected folder).

![Input vs output document folders](~/images/word_folder_diff.png "Input vs output document folders")


## Inserting image in PDF files

Different from the Word files case, for PDF there are two options of reference: relative to text and relative to page, while Word only have relative to text.

### With reference relative to text

In this case, it is pretty similar as the Word case, only changin for both individual files and folder cases, *the Document input type*.

#### For individual files:

1. Select *PDF Document* as document input type.

2. Select *Files* as *Type*.

3. Click on *Explore* in *Files/folder Path* row to select a valid PDF file or list of files. 
    
    ![Options to open PDF file](~/images/pdf_file_1_2_3.png "Options to open PDF file")

     It opens the file explorer  where you should localize the document and open it (you can also select various documents and insert the image in all of them).

    ![Open PDF file(s)](~/images/pdf_file_3.png "Open PDF file(s)")

Continue from 4 to 8 as the [Word](#inserting-image-in-ms-word-files) case for individual files. Make sure that *Relative to text* option is selected in the *Reference type* row.

After the program has completed that task sucessfully, you can check the result PDF in the created subfolder (with name PDF if not changed in the options) under the directory of the original file . 

![Input vs output document files](~/images/pdf_file_diff.png "Input vs output document files")

#### For a folder containing PDF files:

1. Select *PDF Document* as document input type.

2. Select *Folder* as *Type*.

3. Click on *Explore* in *Files/folder Path* row to select the folder containing the PDF files.

    ![Options to open the folder](~/images/pdf_folder_1_2_3.png "Options to open the folder")

     It opens the folder explorer  where you should localize the folder that contains the PDF files.

    <img src="~/images/pdf_folder_3.png" alt="Open folder" width="50%"/>

From here, it continues with the 4 to the 8 steps of individual files. When clicking in the *Generate* button it opens the progress dialog box which shows the status of processing for all of the PDF documents inside the selected folder.

<img src="~/images/pdf_folder_progress.png" alt="Progreso de procesamiento" width="70%"/>

When opening the folder you can check the generated files (each one per input pdf document found in selected folder).

![Input vs output document folders](~/images/pdf_folder_diff.png "Input vs output document folders")


### With reference relative to page

In this case it vary as incorporates its own options to refer the image position relative to the page.

#### For individual files:

1. Select *PDF Document* as document input type.

2. Select *Files* as *Type*.

3. Click on *Explore* in *Files/folder Path* row to select a valid PDF file or list of files. 

4. Click on *Explore* in *Picture Path* row to select a valid image file to insert in the documents selected before. 
    When the file is selected via the file explorer, an image preview is shown in the Preview Box. It also loads the dimensions measured in the selected unit, shown in the *Units* row.

5. You can modify the size of the image to insert and if *Keep aspect ratio* is checked, the other dimension is modified accordingly.
    You can also modify the offset from left (*LeftOffset*) and bottom (*BottomOffset*). You can also enter negative numbers to get an offset to the right and to the bottom. 

    ![Set dimensions](~/images/pdf_file_page_1_to_5.png "Set dimensions")

6. Select the option *Relative to page* in *Reference type* row.

7. Enter the page number and select which corner of the page will be the page reference for the image insertion.

    ![Set page reference](~/images/pdf_file_page_6_7.png "Set page reference")

8. You can modify the subfolder name where the documents will be exported, if desired.

9. Click on generate to create a new PDF document with the image inserted.

    ![All steps](~/images/pdf_file_page_9.png "All steps")

    When the process starts it opens a new dialog which shows in a progress bar the advance of the processing and prints informative messages about the current state of the processing flow.
    
    <img src="~/images/pdf_file_page_progress.png" alt="Progreso de procesamiento" width="70%"/>

Then, you can check the result opening the PDF. Don't forget to close the PDF, if you want to regenerate it, otherwise an error will occur as the software cannot write the file if it is open.

![Input vs output document files](~/images/pdf_file_page_diff.png "Input vs output document files")

#### For a folder containing PDF files:

1. Select *PDF Document* as document input type.

2. Select *Folder* as *Type*.

3. Click on *Explore* in *Files/folder Path* row to select the folder containing the PDF files.

    ![Options to open the folder](~/images/pdf_folder_1_2_3.png "Options to open the folder")

     It opens the folder explorer  where you should localize the folder that contains the PDF files.

    <img src="~/images/pdf_folder_3.png" alt="Open folder" width="50%"/>

From here, it continues with the 4 to the 8 steps of [individual files](#for-individual-files-2). Make sure that the reference is set to *Relative to page* as specified in step 6. When clicking in the *Generate* button it opens the progress dialog box which shows the status of processing for all of the PDF documents inside the selected folder.

<img src="~/images/pdf_folder_page_progress.png" alt="Progreso de procesamiento" width="70%"/>

When opening the folder you can check the generated files (each one per input pdf document found in selected folder).

![Input vs output document folders](~/images/pdf_folder_diff.png "Input vs output document folders")