# docx-csv-mailmerge

Generates .docx files from .csv files using a .docx template with mailmerge fields.

## Installing

[Poetry](https://python-poetry.org/) is used for dependency management and
[pyenv](https://github.com/pyenv/pyenv) to manage python installations. Install dependencies via:

    poetry install

To setup a virtual environment with your local pyenv version run:

    poetry shell

## Usage

Move your `.docx` template and `.csv data file` into the folder of this repo and run either of these:

    python docx-csv-mailmerge.py -t example.docx -c example.csv
    python docx-csv-mailmerge.py --template example.docx --csv example.csv    # alternative

Where the arguments are your word template and the delimiter used in the `csv` file and your data to apply to the word template (in .`csv` format),

You can also use absolute paths to the files, just make sure you add "" around these for escaping characters, for example:

    python docx-csv-mailmerge.py --csv "path/to/data.csv" --template "path/to/template.docx"

If your .csv file uses a delimiter other than ";", use the `--delimiter` option. For help, run

    python docx-csv-mailmerge.py --help

## Demo and preparing your files
This demo uses 'Business Letter' template from Microsoft Office Word. In Word, open your desired document, and add mergefields in place where you like to fill the data programmatically. In the "Insert" ribbon, click on "[ ] Field" (likely next to text box or WordArt) and choose MailMerge --> MergeField.

![Insert Field, Mailmerge, Mergefield](images/1_add_field.png)

In my case, it showed `Error! No bookmark name given`. Right click on the text, "toggle field codes".

![Rightcick, toggle field to show the fieldcode](images/2_toggle_field.png)

You should see something like this `{ MERGEFIELD \* MERGEFORMAT}`. Add a name for this field, like so `{ MERGEFIELD DATE \* MERGEFORMAT}`. This name will have to correspond exactly (case sensitive) to the header in your .csv file. When you right click -- update field, the field should now display something like `<<DATE>>`. Make sure you update this field to the correct styling as you like. You can add more fields following the same approach, or copy paste this field and edit its name using the toggle-field trick.

To try out this template, use - or add some lines to - the example_data.csv. Then, run

     python docx-csv-mailmerge.py example_data.csv example_template.docx ";"
