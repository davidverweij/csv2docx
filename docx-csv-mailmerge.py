from mailmerge import MailMerge
import sys
import csv
import click

@click.command()
@click.option('--template', "-t", "templatepath", required=True, help="Path to the .docx template file to be used")
@click.option('--csv', '-c', 'csvpath', required=True, help="Path to the .csv data file to be used")
@click.option('--delimiter', default=";", help='delimiter used in your csv. Default is \';\'')
def generatedocx(templatepath, csvpath, delimiter):
    try:
        print ("Getting .docx template and .csv data files ...")
        csvfile = open(csvpath, 'rb')
        docx = MailMerge(templatepath)
        csvreader = csv.reader(csvfile, delimiter=str(delimiter))

    except IndexError:
        print ("Error: Insufficient arguments")
    except IOError:
        print ("Error: Files not found")
    except:
        print ("Error: Unexpected")
        raise
    else:
        docx_mergefields = docx.get_merge_fields()
        csv_headers = csvreader.next()
        print ("Found these merge fields in the .docx : ", docx_mergefields)
        print ("Found these header fields in the .csv : ", csv_headers)

        for field in docx_mergefields:                      # see if all fields are accounted for in the .csv header
            if field not in csv_headers:
                print("I can't find the mergefield " +
                      field + " in the .csv file. Please" +
                      "check again and ensure the headers " +
                      "are exactly as the mergfields names. Case sensitive.")
                break
        else:                           # if all fields are accounted for
            print("All merge fields are present in the .csv headers. Generating Word files now...")
            csvdict = csv.DictReader(open(csvpath), delimiter=str(delimiter))                        # have to reopen, because the 'next'

            for counter, row in enumerate(csvdict):
                single_document = {key : row[key] for key in docx_mergefields}       # dictionary comprehension to get mergefields from csv
                docx = MailMerge(sys.argv[2])                                        # get the template
                try:
                    docx.merge_pages([single_document])                              # passing it as a 'list' (of one) to allow prepoluated objects as data
                except ValueError:
                    print ("ValueError in document number " + str(counter) + ". Please check the .csv for valid characters. Continuing with the rest...")
                except:
                    print("Uknown Error")
                    raise
                else:
                    docx.write(str(counter) + ".docx")                               # TODO: haven't found a way to write to a subfolder

    print("Finished Program.")

if __name__ == "__main__":
    generatedocx()
