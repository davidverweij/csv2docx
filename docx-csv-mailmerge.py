from mailmerge import MailMerge
import csv
import sys

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
            # TODO: haven't found a way to write to a subfolder
            docx.write(f"{counter}.docx")

if __name__ == '__main__':
    # TODO: validate user input independent of the convert method
    # Showing how unpacking data works
    _, data, template, delimiter = sys.argv
    # Invoke our main function above (ideally via a CLI like click)
    convert(data, template, delimiter)
