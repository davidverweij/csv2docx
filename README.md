# docx-csv-mailmerge
Generates .docx files from .csv files using a .docx template with mailmerge fields.

## Required installed packages
[docx-mailmerge](https://pypi.org/project/docx-mailmerge/)

**NOTE** docx-mailmerge is based on [lxml](https://lxml.de/installation.html), which for MacOS users _like myself_ is currently only available through a MacPort (see also Bonus below). This then only allows the use of Python 2.7 - not 3. That is why this code is executed with Python 2.7. Running it on 3 will needs some tweaks (particularly on the print statements).

## Usage
If you have the required package(s) installed (see above), put the docx-csv-mailmerge.py into the same folder as your .docx template and .csv data file. Then, run:

    python docx-csv-mailmerge.py data.csv template.docx ";"

where the first argument is the .csv file, the second the template, and the third the delimiter as used in the .csv between "quotation marks". You can also use absolute paths to the files, just make sure you add "" around these for escaping characters, for example:

    python docx-csv-mailmerge.py "path/to/data.csv" "path/to/template.docx" "/"

## Demo and preparing your files
This demo uses 'Business Letter' template from Microsoft Office Word. In Word, open your desired document, and add mergefields in place where you like to fill the data programmatically. In the "Insert" ribbon, click on "[ ] Field" (likely next to text box or WordArt) and choose MailMerge --> MergeField.

![Insert Field, Mailmerge, Mergefield](images/1_add_field.png)

In my case, it showed `Error! No bookmark name given`. Right click on the text, "toggle field codes".

![Insert Field, Mailmerge, Mergefield](images/1_add_field.png)

You should see something like this `{ MERGEFIELD \* MERGEFORMAT}`. Add a name for this field, like so `{ MERGEFIELD DATE \* MERGEFORMAT}`. This name will have to correspond exactly (case sensitive) to the header in your .csv file. When you right click -- update field, the field should now display something like `<<DATE>>`. Make sure you update this field to the correct styling as you like. You can add more fields following the same approach, or copy paste this field and edit its name using the toggle-field trick.

To try out this template, use - or add some lines to - the example_data.csv. Then, run

     python docx-csv-mailmerge.py example_data.csv example_template.docx ";"

## Bonus
Use [docx2pdf](https://github.com/AlJohri/docx2pdf) to batch convert the generated .docx documents into .pdfs. Mac users need to use Python 2.7 and [MacPorts](https://www.macports.org/install.php) to install this package. In my experience, I had to run docx2pdf once, grant access to Microsoft Word, and could then run the whole batch nicely. I first moved all generated .docx to a folder to use the folder-batch from docx2pdf easily.
