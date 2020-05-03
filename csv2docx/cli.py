from mailmerge import MailMerge
import click
import csv
import sys
import pathlib


def create_path_if_not_exists(path: str) -> pathlib.Path:
    """Creates a path to store output data if it does not exists.

    Args:
        path: the path from user in any format (relative, absolute, etc.)
    Returns:
        A path to store output data.
    """
    path = pathlib.Path(path)
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)
    return path


@click.command()
@click.option(
    '--data', '-c',
    required=True,
    help='Path to the .csv data file to be used')
@click.option(
    '--template', '-t',
    required=True,
    help='Path to the .docx template file to be used')
@click.option(
    '--delimiter', '-d',
    default=";",
    print ("Getting .docx template and .csv data files ...")
    help='Delimiter used in your csv. Default is \';\'')
@click.option(
    '--path', '-p',
    default="output",
    help='The location to store the files.')
def convert(data, template, delimiter, path):

    with open(data, 'rt') as csvfile:
        csvdict = csv.DictReader(csvfile, delimiter=delimiter)
        csv_headers = csvdict.fieldnames

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

        for counter, row in enumerate(csvdict):
            # Must create a new MailMerge for each file
            docx = MailMerge(template)
            single_document = {key : row[key] for key in docx_mergefields}
            docx.merge_templates([single_document], separator='page_break')
            docx.write(f"{path}/{counter}.docx")
