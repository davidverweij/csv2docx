import csv
from mailmerge import MailMerge
from pathlib import Path

def create_path_if_not_exists(path: str):
    """Creates a path to store output data if it does not exists.
    Args:
        path: the path from user in any format (relative, absolute, etc.)
    Returns:
        A path to store output data.
    """
    path = Path(path)
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)
    return path

def generate_name(filename, path):
    """Checks whether file with same name exists or not
    Args:
        filename: the file we want to check
        path: the directory we want to check in
    Returns:
        An unique full path with directory.
    """
    filename += ".docx"
    checkpath = path / filename
    if checkpath.exists():
        # Count available files with same name
        counter = len(list(path.glob(checkpath.stem + "*docx" ))) + 1
        return f"{path}/{checkpath.stem}_{counter}.docx"
    return checkpath

def convert(data, template, delimiter=";", custom_name=None, path="output"):
    print ("Getting .docx template and .csv data files ...")

    with open(data, 'rt') as csvfile:
        csvdict = csv.DictReader(csvfile, delimiter=delimiter)
        csv_headers = csvdict.fieldnames
        if (custom_name != None and custom_name not in csv_headers):
                print("Column name not found. Please enter valid column name")
                exit()
        docx = MailMerge(template)
        docx_mergefields = docx.get_merge_fields()

        print(f"DOCX fields : {docx_mergefields}")
        print(f"CSV field   : {csv_headers}")

        # see if all fields are accounted for in the .csv header
        column_in_data = set(docx_mergefields) - set(csv_headers)
        if len(column_in_data) > 0:
            print (f"{column_in_data} is in the word document, but not csv.")
            return

        print("All fields are present in your csv. Generating Word docs ...")
        path = create_path_if_not_exists(path)

        for row in csvdict:
            # Must create a new MailMerge for each file
            docx = MailMerge(template)
            single_document = {key : row[key] for key in docx_mergefields}
            docx.merge_templates([single_document], separator='page_break')
            # Striping every name to remove extra spaces
            filename = generate_name(row[custom_name].strip(), path)
            docx.write(filename)