import csv
from mailmerge import MailMerge
from typing import List, Optional


def unique_filenames(
    csv_path: str,
    column_name: str,
    delimiter: str
) -> List[str]:
    filenames: List[str] = []
    with open(csv_path, "rt") as csvfile:
        csvdict = csv.DictReader(csvfile, delimiter=delimiter)
        column = [row[column_name] for row in csvdict]

        for index, item in enumerate(column):
            # Count items up to but excluded this item in the column
            num_dups_from_pos = column[:index].count(item)
            # Only change item if its a duplicate
            if num_dups_from_pos > 1:
                item = f"{item}_{num_dups_from_pos}"
            filenames.append(item)
    return filenames


def convert(
    data: str,
    template: str,
    delimiter: Optional[str] = ";",
    column_name: Optional[str] = None
) -> None:
    print("Getting .docx template and .csv data files ...")

    with open(data, "rt") as csvfile:
        csvdict = csv.DictReader(csvfile, delimiter=delimiter)
        csv_headers = csvdict.fieldnames

        if column_name and column_name not in csv_headers:
            print("Column name not found. Please enter valid column name.")
            exit()
        elif column_name:
            filenames = unique_filenames(data, column_name, delimiter)

        docx = MailMerge(template)
        docx_mergefields = docx.get_merge_fields()

        print(f"DOCX fields : {docx_mergefields}")
        print(f"CSV field   : {csv_headers}")

        # see if all fields are accounted for in the .csv header
        column_in_data = set(docx_mergefields) - set(csv_headers)

        if len(column_in_data) > 0:
            print(f"{column_in_data} is in the word document, but not csv.")
            return

        print("All fields are present in your csv. Generating Word docs ...")

        for index, row in enumerate(csvdict):
            docx = MailMerge(template)
            single_document = {key: row[key] for key in docx_mergefields}
            docx.merge_templates([single_document], separator="page_break")
            filename = filenames[index] if column_name else index
            docx.write(f"{filename}.docx")
