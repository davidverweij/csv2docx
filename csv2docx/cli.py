from mailmerge import MailMerge
import click
import csv
import sys

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
    help='delimiter used in your csv. Default is \';\'')
def convert(data, template, delimiter):
    print ("Getting .docx template and .csv data files ...")

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

        for counter, row in enumerate(csvdict):
            # Must create a new MailMerge for each file
            docx = MailMerge(template)
            single_document = {key : row[key] for key in docx_mergefields}
            docx.merge_templates([single_document], separator='page_break')
            # TODO: write to user-defined subfolder
            docx.write(f"{counter}.docx")
