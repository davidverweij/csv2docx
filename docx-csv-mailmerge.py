#!/usr/bin/env python

__author__ = "David Verweij"
__license__ = "MIT"
__version__ = "0.1"
__email__ = "davidverweij@gmail.com"
__status__ = "Work in Progress"

from mailmerge import MailMerge
import sys
import csv

if __name__ == '__main__':                      #code to execute if called from command-line
    try:
        print ("Getting .docx template and .csv data files ...")
        csvfile = open(sys.argv[1], 'rb')
        docx = MailMerge(sys.argv[2])
        delimiter = sys.argv[3]
        csvreader = csv.reader(csvfile, delimiter=delimiter)

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
            csvdict = csv.DictReader(open(sys.argv[1], 'rb'), delimiter=delimiter)                        # have to reopen, because the 'next'

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
