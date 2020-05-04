import csv
from mailmerge import MailMerge

def generate_names(listnm):
    newname = []
    for i in range(len(listnm)):
        if (listnm[i] not in listnm[:i]):
            newname.append(listnm[i])
        else:
            newname.append(listnm[i] + "_" + str(listnm[:i].count(listnm[i]) + 1))
    return newname

def convert(data, template, delimiter=";", custom_name="NULL"):
    print ("Getting .docx template and .csv data files ...")

    with open(data, 'rt') as csvfile:
        csvdict = csv.DictReader(csvfile, delimiter=delimiter)
        csv_headers = csvdict.fieldnames
        if (custom_name != None):
            if (custom_name not in csv_headers):
                print("Column name not found. Please enter valid column name")
                exit()
            else:
                file_names = generate_names(list(row[custom_name] for row in csvdict))
                csvfile.seek(1)
                csvfile.readline()
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

            docx.write(f"{file_names[counter] if (custom_name != None) else counter+1}.docx")          