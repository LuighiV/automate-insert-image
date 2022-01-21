# Use of the software as command line application

This is the alternative mode of use for the software. It is useful when trying to use as part of a script. For use it, you should locate the executable, which should be under `C:\Program Files\InsertImageDocument`, and has the name `editdocuments.exe`. If you want, you can add this folder to the PATH environment variable or simply execute the file from this location.

## Usage

You should use the executable name followed by any of the following options. If no option is entered, the sofware will be lauch the graphical interface.
```
editdocuments.exe --option_1 value_1 --option_2 value_2 ... --option_n value_n
```

Where `option` could be the following:

*  `-t, --type`              (Default: `Word`) Document type. Valid values: `Word`, `PDF`.

*  `-A, --is-absolute`       (Default: `false`) The reference is absolute to the page. When it is true uses a specific page and a corner in the page as reference (only for PDF type).

*  `-N, --page`              (Default: 1) Page number (used when `--is-absolute` is `true`).

*  `-r, --page-reference`    (Default: `bottom_left`) Page reference (used when `--is-absolute` is `true`). Valid values: `bottom_left`, `top_left`, `top_right`, `bottom_right`.

*  `-f, --inputfile`          Input files to be processed.

*  `-d, --inputfolder`       Input folder to be processed.

*  `-p, --picturepath`       Picture file to be inserted.

*  `-H, --placeholder`       Placeholder where the picture is inserted.

*  `--verbose`               (Default: `false`) Prints all messages to standard output.

*  `-V, --visible`           (Default: `false`) Set MSWord Application visible (only for Word type).

*  `-S, --foldersave`        Set folder to save PDF.

*  `-u, --subfoldersave`     Save in subfolder relative to document path

*  `-s, --save`              (Default: `false`) Save in file.

*  `--leftoffset`            (Default: 0) Left offset for picture (in points).

*  `--bottomoffset`          (Default: 0) Bottom offset for picture (in points).

*  `--width`                 (Default: 200) Width for picture (in points).

*  `--height`                (Default: 150) Height for picture (in points).

*  `--help`                  Display this help screen.

*  `--version`               Display version information.

Keep in mind that the equivalence for points is 72 points = 1 inch = 2.54 cm.

## Examples

There ara various cases of use for the software. In the following exampels we show the most common:

### For a single Word file

This command receives the Word file as an input and insert the image in a position around the placeholder text. Its counterpart in the GUI mode is [Inserting image in MS Word files (for individual files)](gui.md#for-individual-files)

```bash
.\editdocuments.exe --inputfile "D:\Luighi\Documents\test-iid\document_en.docx" --picturepath "D:\Luighi\Documents\test-iid\signature_hw.png" --placeholder "signature" --subfoldersave "PDF" --bottomoffset "-7" --leftoffset "-40" --width "140" --height "56" --verbose true
```

### For a folder with Word files

This command insert the image to all the Word files in the specified input folder.  Its counterpart in the GUI mode is [Inserting image in MS Word files (for a folder containing Word files)](gui.md#for-a-folder-containing-word-files)

```bash
.\editdocuments.exe --inputfolder "D:\Luighi\Documents\test-iid\test_folder" --picturepath "D:\Luighi\Documents\test-iid\signature_hw.png" --placeholder "signature" --subfoldersave "PDF" --bottomoffset "-7" --leftoffset "-40" --width "140" --height "56" --verbose true
```


### For a folder with PDF files (relative to text)

In this case, the image is inserted in all the PDF files in the specified input folder.  Its counterpart in the GUI mode is [Inserting image in PDF files (relative to text - for a folder containing PDF files)](gui.md#for-a-folder-containing-pdf-files)

```bash
.\editdocuments.exe --type PDF --inputfolder "D:\Luighi\Documents\test-iid\test_folder_pdf" --picturepath "D:\Luighi\Documents\test-iid\signature_hw.png" --placeholder "signature" --subfoldersave "PDF" --bottomoffset "-7" --leftoffset "-40" --width "140" --height "56" --verbose true
```

### For a folder with PDF files (relative to page)

This command insert the image with a position relative to the page instead of the text. Its counterpart in the GUI mode is [Inserting image in MS Word files (relative to page - for a folder containing PDF files)](gui.md#for-a-folder-containing-pdf-files-1)

```bash
.\editdocuments.exe --type PDF --inputfolder "D:\Luighi\Documents\test-iid\test_folder_pdf" --picturepath "D:\Luighi\Documents\test-iid\signature_hw.png" --is-absolute true --page 1 --page-reference top_left --subfoldersave "PDF" --bottomoffset "-80" --leftoffset "57" --width "85" --height "71" --verbose true
```